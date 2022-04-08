Attribute VB_Name = "VBT_LIB_Common"
Option Explicit
'Revision History:
'V0.0 initial bring up
'V0.1 add keep alive function
'V0.2 add disable compare and enable compare function.
'variable declaration
Public Const Version_Lib_VBT_Common = "0.1"  'lib version

Public DicDiffPairs As New Scripting.Dictionary  'relocation for minimum VBT with RF code'*****************************************
'*****************************************
'******               Relay controls******
'*****************************************

Public Function Relay_Control(Optional relay_on As PinList, Optional relay_off As PinList, Optional WaitTime As Double = 0.003)
'control relay on off, will auto trim NC pins Tto cover CP, FT both stages
    Dim Pins_On() As String, Pin_Cnt_On As Long
    Dim Pins_Off() As String, Pin_Cnt_Off As Long
    Dim p As Variant
    Dim relayOnStr As String, relayOffStr As String
    Dim wait_time As Double 'relay wiat time by global spec
    Dim Tname As String
    Dim BitState As New PinListData
    Dim i, j, k, relay_off_gpnum, relay_off_divnum As Long

    On Error GoTo errHandler
    
    relayOnStr = ""
    relayOffStr = ""
    Tname = ""
    TheExec.DataManager.DecomposePinList relay_on, Pins_On(), Pin_Cnt_On
    TheExec.DataManager.DecomposePinList relay_off, Pins_Off(), Pin_Cnt_Off

        relay_off_divnum = 10
    If Pin_Cnt_Off > relay_off_divnum Then
        relay_off_gpnum = Pin_Cnt_Off \ relay_off_divnum
        ReDim relay_off_arr(relay_off_gpnum)
        For i = 0 To relay_off_gpnum
            k = relay_off_divnum * i
            For j = k To (relay_off_divnum - 1) + k
                If j = Pin_Cnt_Off Then Exit For
                If relay_off_arr(i) = "" Then
                    relay_off_arr(i) = relay_off_arr(i) & Pins_Off(j)
                Else
                    relay_off_arr(i) = relay_off_arr(i) & ", " & Pins_Off(j)
                End If
            Next j
        Next i
    End If

    Trim_NC_Pin Pins_On, Pin_Cnt_On
    Trim_NC_Pin Pins_Off, Pin_Cnt_Off
    
    If Pin_Cnt_On <> 0 Then
        TheHdw.Utility.Pins(relay_on).State = tlUtilBitOn
        For Each p In Pins_On
            Tname = "rly_on_" & p
            BitState = TheHdw.Utility.Pins(p).States(tlUBStateProgrammed)
            If relayOnStr = "" Then
                relayOnStr = relayOnStr & p
            Else
                relayOnStr = relayOnStr & ", " & p
            End If
        'TheExec.Datalog.WriteComment "Relay On : " & relayOnStr
        'TheExec.Flow.TestLimit resultVal:=BitState.Pins(p), lowval:=tlUtilBitOn, hival:=tlUtilBitOn, Tname:=Tname, ForceResults:=tlForceNone
        Next p
        TheExec.Datalog.WriteComment "Relay On : " & relayOnStr
    End If
    
    If Pin_Cnt_Off <> 0 Then
        ' 20181226 prevent DIB:0004 alarm
        If Pin_Cnt_Off > relay_off_divnum Then
            For i = 0 To relay_off_gpnum
                TheHdw.Utility.Pins(relay_off_arr(i)).State = tlUtilBitOff
                TheHdw.Wait 0.01
            Next i
        Else
            TheHdw.Utility.Pins(relay_off).State = tlUtilBitOff
        End If
        
        For Each p In Pins_Off
            Tname = "rly_off_" & p
            BitState = TheHdw.Utility.Pins(p).States(tlUBStateProgrammed)
            If relayOffStr = "" Then
                relayOffStr = relayOffStr & p
            Else
                relayOffStr = relayOffStr & ", " & p
            End If
        'TheExec.Datalog.WriteComment "Relay off : " & relayOffStr
        'TheExec.Flow.TestLimit resultVal:=BitState.Pins(p), lowval:=tlUtilBitOff, hival:=tlUtilBitOff, Tname:=Tname, ForceResults:=tlForceNone
        Next p
        TheExec.Datalog.WriteComment "Relay off : " & relayOffStr
    End If
    
    Wait WaitTime
    
    Exit Function
errHandler:
    ErrorDescription ("Relay_Control")
    If AbortTest Then Exit Function Else Resume Next
End Function

'*****************************************
'******         free run clk, nWire ******
'*****************************************


Public Function StartSBClock(SBFreq As Double) As Long
    On Error GoTo errHandler
    Dim SBC_Enable As Long
    'Dim SBFreq As Double
    'TheExec.Datalog.WriteComment "******************  Enable Support BD clock ****************"
    'SBFreq = TheExec.specs.Globals("SBC_Freq_Var").ContextValue
    
    With TheHdw.DIB.SupportBoardClock
        .Connect
        .Frequency = SBFreq
        .Vih = XI0_ref_VOH ' Max is 6V
        .Vil = 0 ' Min is -1V
        .start
    End With
    SBC_Enable = 1
    TheExec.Flow.TestLimit SBC_Enable, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="SBC enable" 'BurstResult=1:Pass
    'printing in data log
    TheExec.Datalog.WriteComment "********** support board clock = " & Format(SBFreq / 1000000, "0.000") & " Mhz, Clock_Vih = " _
                             & XI0_ref_VOH & " V, Clock_Vil = " & XI0_ref_VOL & " V  *******"
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function StopSBClock()
    Dim SBC_Enable As Long
    ' Stop and disconnect the support board clock.
    With TheHdw.DIB.SupportBoardClock
        .stop
        .Disconnect
    End With
    SBC_Enable = 0
    TheExec.Flow.TestLimit SBC_Enable, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="SBC disable" 'BurstResult=1:Pass
    'printing in data log
    TheExec.Datalog.WriteComment "******************  Disable Support BD clock ****************"
End Function
Public Function FreeRunclk_Enable_ori(PortName As String)

    Dim site As Variant
    Dim i As Long
    Dim PLLLockChecked As New SiteLong
    Dim measf As New PinListData
    
    Dim NotLocked As Boolean
    Dim XI0_REFCLK As String
    Dim PortMode As String
    Dim PLL_Lock As New SiteLong
    Dim port_level_name As String
    Dim port_level_value As Double
    Dim FreeRunFreq As Double
    On Error GoTo errHandler
    
    'Default to stop nWire before apply new spec
    'If TheExec.Flow.EnableWord("XI0_nWire") = True Then ' remove this enable word, must trigger N-wire
        TheHdw.Protocol.ports(PortName).Halt
        TheHdw.Protocol.ports(PortName).Enabled = False
    'End If
    
    'CurrentXi0Freq = FreeRunFreq
    'If TheExec.Flow.EnableWord("XI0_nWire") = True Then  remove this enable word, must trigger N-wire
        If TheExec.DataManager.instanceName <> "" Then
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        Else
            TheHdw.Digital.ApplyLevelsTiming False, True, True, tlPowered, , , , Level_nWire, "common", "typ", TSB_nWire, "XI0_24M_TCK_24M", "typ"
        End If

        
        '' 20151028 - Setup FRCPath by input port name
        TheHdw.Protocol.ports(PortName).Enabled = True
        TheHdw.Protocol.ports(PortName).NWire.ResetPLL
        TheHdw.Wait 0.001
        
        '///////////////////////////////////////////////////////////////////////
        If LCase(PortName) Like "*diff*" Then
               port_level_name = Replace(LCase(PortName), "port", "pa")
               port_level_value = TheHdw.Digital.Pins(port_level_name).DifferentialLevels.Value(chVid)
        Else
               port_level_name = Replace(LCase(PortName), "port", "pa")
               port_level_value = TheHdw.Digital.Pins(port_level_name).Levels.Value(chVih)
        End If
        '///////////////////////////////////////////////////////////////////////
        
        ' Test the PLL lock status and datalog results.
        For Each site In TheExec.sites.Selected
            If TheHdw.Protocol.ports(PortName).NWire.IsPLLLocked = False Then
                PLLLockChecked = 1
                'TheExec.Datalog.WriteComment "print: site(" & Site & "), PortName(" & PortName & "), IsPLLLocked = False"
                PLL_Lock(site) = 0
            Else
                PLLLockChecked = 0
                'TheExec.Datalog.WriteComment "print: site(" & Site & "), PortName(" & PortName & "), IsPLLLocked = True"   'comment only
                'TheExec.Datalog.WriteComment "print: site(" & Site & "), PortName(" & PortName & "), Wake up finished"
                PLL_Lock(site) = 1
            End If
        Next site
        
        ' Start the nWire engine.
        Call TheHdw.Protocol.ports(PortName).NWire.Frames("RunFreeClock").Execute
        
        TheHdw.Protocol.ports(PortName).IdleWait
        TheHdw.Wait 0.001
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites.Selected
                PLL_Lock(site) = 1
            Next site
        End If

        TheExec.Flow.TestLimit PLL_Lock, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="nWirePLL_Lock" 'BurstResult=1:Pass

        'TheExec.Datalog.WriteComment ""
        'TheExec.Datalog.WriteComment "********** Enable freerunning clock *********"
                      
        '****print out to data log about nWire clock condition
        FreeRunFreq_debug = 1 / TheHdw.Digital.Timing.Period(PortName) / 1000000

        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            FreeRunFreq = 24000000  'offline
            FreeRunFreq_debug = FreeRunFreq / 1000000
        End If
        
        TheExec.Datalog.WriteComment "********** freerunning clock = " & Format(FreeRunFreq_debug, "0.000") & " Mhz, port_level_value" = " & port_level_value"

    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function
Public Function FreeRunClk_Disable(PortName As String, Optional powerdown_flag As Boolean = False) As Long
    Dim site As Variant
    
    Call Disable_FRC(PortName)  '''''Support multiple nWire port 20170718'''''''''''''
''    ' Disable the nWire engine.
''    'If TheExec.Flow.EnableWord("XI0_nWire") = True Then removed
''        TheHdw.Protocol.ports(PortName).Halt
''        TheHdw.Protocol.ports(PortName).Enabled = False     'scope out point
''    'End If
    If powerdown_flag = False Then TheExec.Flow.TestLimit 0, 0, 0, tlSignGreaterEqual, tlSignLessEqual, Tname:="nWire halt" 'BurstResult=1:Pass
    'printing to data log
    'TheExec.Datalog.WriteComment "******************  Disable freerunning clock ****************"
    
    ''upload to global constant
    FreeRunFreq_debug = 0
    clock_Vih_debug = 0
    clock_Vil_debug = 0

End Function

Function Start_Profile(PinName As PinList, WhatToCapture As String, SampleRate As Double, SampleSize As Long, Optional CapSignalName As String = "Capture_signal")
'start current or voltage profile capturing

On Error GoTo errHandler
' Wait if another capture is running
    Do While TheHdw.DCVS.Pins(PinName).Capture.IsRunning = True
    Loop
    
    'Create a SIGNAL to set up instrument
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Add CapSignalName
    
    'Set this as the default signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.DefaultSignal = CapSignalName
    
    'Define the signal used for the capture
    With TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName)
        .Reinitialize
        If (WhatToCapture = "I") Then
            .mode = tlDCVSMeterCurrent
            .range = TheHdw.DCVS.Pins(PinName).CurrentRange.max '2
        Else
            .mode = tlDCVSMeterVoltage
            .range = 10
        End If
        .SampleRate = SampleRate
        .SampleSize = SampleSize
    
    End With
    
    ' Setup the hardware by loading the signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName).LoadSettings
    
    ' Start the capture
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName).Trigger

    Exit Function
errHandler:
    ErrorDescription ("Start_Profile")
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function start_profile_DCVI(PinName As String, WhatToCapture As String, SampleRate As Double, SampleSize As Long)

Do While TheHdw.DCVI.Pins(PinName).Capture.IsCaptureDone = False        ' Wait if another capture is running
Loop
TheHdw.DCVI.Pins(PinName).Capture.Signals.Add "Capture_signal"              'Create a SIGNAL to set up instrument
TheHdw.DCVI.Pins(PinName).Capture.Signals.DefaultSignal = "Capture_signal"  'Set this as the default signal

        With TheHdw.DCVI.Pins(PinName)
            .Gate = False
            .mode = tlDCVIModeCurrent
            .Voltage = 6
            .VoltageRange.Autorange = True
            .CurrentRange.Autorange = True
            .current = 0
            .Connect tlDCVIConnectDefault
            .Gate = True
        End With
 
With TheHdw.DCVI.Pins(PinName).Capture.Signals.Item("Capture_signal")    ' Define the signal used for the capture
    .Reinitialize
    If (WhatToCapture = "I") Then
        .mode = tlDCVIMeterCurrent
        .range = 0.02
    Else
        .mode = tlDCVIMeterVoltage
         .range = 7
    End If
    .SampleRate = SampleRate
    .SampleSize = SampleSize
End With

TheHdw.DCVI.Pins(PinName).Capture.Signals.Item("Capture_signal").LoadSettings  ' Setup the hardware by loading the signal
TheHdw.DCVI.Pins(PinName).Capture.Signals.Item("Capture_signal").Trigger            ' Start the capture
End Function


Public Function Plot_Profile(PinName As PinList, Optional CapSignalName As String = "Capture_signal", Optional ExportWaveform As Boolean = False)
'Plot profiles

    Dim DSPW As New DSPWave
    Dim Label As String
    Dim site As Variant
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim p As Variant
    Dim FileName As String
    Dim lastBurstPat As New SiteVariant
    Dim isGrp As New SiteBoolean
    Dim lastLabel As New SiteVariant
    Dim day_code As String
    Dim Current_Insatance As String
    Current_Insatance = m1_InstanceName
    If Current_Insatance = "" Then TheExec.Datalog.WriteComment "<ERROR> Instance name is empty.Please check instance global name is defined."
    
    On Error GoTo errHandler

    Do While TheHdw.DCVS.Pins(PinName).Capture.IsRunning = True
    Loop

    day_code = CStr(Year(Now)) & Right("0" & CStr(Month(Now)), 2) & Right("0" & CStr(day(Now)), 2)
    day_code = day_code & Right("0" & CStr(Hour(Now)), 2) & Right("0" & CStr(Minute(Now)), 2) & Right("0" & CStr(Second(Now)), 2)
    ' Get the captured samples from the instrument
    Call TheExec.DataManager.DecomposePinList(PinName, Pin_Ary(), Pin_Cnt)
    Dim sampleR As String
    For Each p In Pin_Ary
        If TheExec.DataManager.ChannelType(p) <> "N/C" Then
            DSPW = TheHdw.DCVS.Pins(p).Capture.Signals(CapSignalName).DSPWave
            For Each site In TheExec.sites
'                TheHdw.Digital.Patgen.ReadLastStart lastBurstPat, isGrp, lastLabel
                 sampleR = CStr(TheHdw.DCVS.Pins(p).Capture.SampleRate)
                'If thehdw.DCVS.Pins(p).Meter.mode = thehdw.DCVS.Pins(p).CurrentRange.Max Then
                If TheHdw.DCVS.Pins(p).Meter.mode = tlDCVSMeterCurrent Then
                    Label = "Current Profile for Site: " & site & " " & " " & CapSignalName & "Pin :" & " " & p
                    FileName = "CurrentProfile-Site" & site & "-" & p & "-" & sampleR & "-" & Current_Insatance & "_" & day_code & ".txt"
                Else
                    Label = "Voltage Profile for Site: " & site & " " & " " & CapSignalName & "Pin :" & " " & p
                    FileName = "VoltageProfile-Site" & site & "-" & p & "-" & sampleR & "-" & Current_Insatance & "_" & day_code & ".txt"
                End If
                
                If True Then DSPW.plot Label   'for pliot
                If ExportWaveform Then
                    Dim TempStr As String
                    TempStr = "D:\" & p
                    Dim fso As New FileSystemObject
                     
                    If Dir(TempStr, vbDirectory) = Empty Then
                        MkDir TempStr
                    End If
                    DSPW.FileExport TempStr & "\" & FileName, File_txt
                End If
                'If LCase(GetInstrument(CStr(p), 0)) <> "hexvs" Then DSPW.Clear
                Set DSPW = Nothing
            Next site
        End If
    Next p
    m1_InstanceName = ""
    Exit Function
errHandler:
    ErrorDescription ("Plot_Profile")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Plot_profile_DCVI(PinName As String)

Dim DSPW As New DSPWave
Dim Label As String
Dim site As Variant

On Error GoTo errHandler

' Get the captured samples from the instrument
DSPW = TheHdw.DCVI.Pins(PinName).Capture.Signals.Item("Capture_signal").DSPWave

For Each site In TheExec.sites.Active
    If TheHdw.DCVI.Pins(PinName).Meter.mode = tlDCVIMeterCurrent Then
        Label = "Current Profile for Site: " & site
    Else
        Label = "Voltage Profile for Site: " & site
End If
    
     DSPW.plot Label
    
Next site

Exit Function

errHandler:
        TheExec.AddOutput "Error in the Plot Profile"
                If AbortTest Then Exit Function Else Resume Next
End Function

Function Start_Profile_AutoResolution(PinName As String, WhatToCapture As String, SampleRate As Double, SampleSize As Long, Optional CapSignalName As String = "Capture_signal", Optional Plottime As Double = 0)
'start current or voltage profile capturing

On Error GoTo errHandler
    Dim Profile_AllPin() As String
    Dim PinCnt As Long
    Dim HexPins As String
    Dim UVSPins As String
    Dim Pin As Variant
    Dim Profile_SampleRate_Hex As Double
    Dim Profile_SampleSize_Hex As Double

    Dim Profile_SampleRate_UVS As Double
    Dim Profile_SampleSize_UVS As Double
   
    If Plottime <> 0 Then
        
        SplitPinByinstrument PinName, HexPins, UVSPins
        
        If HexPins <> "" Then
            Call ProfileAutoResolution("HEX", Plottime, Profile_SampleSize_Hex, Profile_SampleRate_Hex)
            StartProfile HexPins, WhatToCapture, Profile_SampleRate_Hex, Profile_SampleSize_Hex, CapSignalName, "HEX"
        End If
        If UVSPins <> "" Then
            Call ProfileAutoResolution("UVS", Plottime, Profile_SampleSize_UVS, Profile_SampleRate_UVS)
            StartProfile UVSPins, WhatToCapture, Profile_SampleRate_UVS, Profile_SampleSize_UVS, CapSignalName, "UVS"
        End If
        Exit Function
    Else
        ' Wait if another capture is running
        Do While TheHdw.DCVS.Pins(PinName).Capture.IsRunning = True
        Loop
        
        'Create a SIGNAL to set up instrument
        TheHdw.DCVS.Pins(PinName).Capture.Signals.Add CapSignalName
        
        'Set this as the default signal
        TheHdw.DCVS.Pins(PinName).Capture.Signals.DefaultSignal = CapSignalName
        
        'Define the signal used for the capture
        With TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName)
            .Reinitialize
            If (UCase(WhatToCapture) = "I") Then
                .mode = tlDCVSMeterCurrent
                .range = TheHdw.DCVS.Pins(PinName).CurrentRange.max '2
            Else
                .mode = tlDCVSMeterVoltage
                .range = 10
            End If
            .SampleRate = SampleRate
            .SampleSize = SampleSize
        
        End With
        
        ' Setup the hardware by loading the signal
        TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName).LoadSettings
        
        ' Start the capture
        TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName).Trigger
    
        Exit Function
    End If

errHandler:
    ErrorDescription ("Start_Profile_AutoResolution")
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Print_Footer(PrintInfo As String)

    TheExec.Datalog.WriteComment "******************************"
    TheExec.Datalog.WriteComment "*print: " & PrintInfo & " end*"
    TheExec.Datalog.WriteComment "******************************"

End Function
Public Function Print_Header(PrintInfo As String)

    TheExec.Datalog.WriteComment "********************************"
    TheExec.Datalog.WriteComment "*print: " & PrintInfo & " start*"
    TheExec.Datalog.WriteComment "********************************"

End Function
'Public Function Start_Current_Profile(PinName As PinList, SampleRate As Double, SampleSize As Long)
'    TheExec.EnableWord("Profile_Voltage") = False
'    Call Start_Profile(PinName, "I", SampleRate, SampleSize)
'End Function
'
'Public Function Start_Voltage_Profile(PinName As PinList, SampleRate As Double, SampleSize As Long)
'    TheExec.EnableWord("Profile_Current") = False
'    Call Start_Profile(PinName, "V", SampleRate, SampleSize)
'End Function

Public Function Print_PgmInfo(separate As Boolean)
    Dim PatVersion As String
    Dim Excel_Version_SOC_SCAN As String
    Dim Excel_Version_SOC_MBIST As String
    Dim Excel_Version_CPU_SCAN As String
    Dim Excel_Version_CPU_MBIST As String
    Dim Excel_Version_CPU_PCM As String
    Dim Excel_Version_HARDIP As String
    Dim Excel_Version As String
    Dim ws_pat As Worksheet
    Dim wb As Workbook
    Dim ws_pat1 As Worksheet
    Dim wb1 As Workbook
    Dim TestPlanVersion As String
    
    Set wb = Application.ActiveWorkbook
    Set ws_pat = wb.Sheets("PatSets_ALL")
    
    PatVersion = ws_pat.Cells(4, 10).Value
    Excel_Version = ws_pat.Cells(5, 10).Value
    TestPlanVersion = ws_pat.Cells(6, 10).Value
   
    ''Rhea_A0_Patlist_CPU_PCM#4_EXT_140917
    ''Rhea_A0_Patlist_CPU_MBIST#11_EXT_140916
    ''Rhea_A0_Patlist_CPU_SCAN#4_EXT_140902
    ''Rhea_A0_Patlist_SOC_MBIST#5_EXT_140915
    ''Rhea_A0_Patlist_SOC_SCAN#7_EXT_140909
    ''Rhea_A0_Vector_List_HardIP#88_EXT_140922
    'V01B
'''''''    TestPlanVersion = "Cayman_ATE_TestPlan#2"
'''''''
'''''''    If separate = True Then
'''''''        Excel_Version_SOC_SCAN = "Rhea_A0_Patlist_SOC_SCAN#8_EXT_140926"
'''''''        Excel_Version_SOC_MBIST = "Rhea_A0_Patlist_SOC_MBIST#5_EXT_140915"
'''''''        Excel_Version_CPU_SCAN = "Rhea_A0_Patlist_CPU_SCAN#7_EXT_141002"
'''''''        Excel_Version_CPU_MBIST = "Rhea_A0_Patlist_CPU_MBIST#13_EXT_140930"
'''''''        Excel_Version_CPU_PCM = "Rhea_A0_Patlist_CPU_PCM#4_EXT_140917"
'''''''        Excel_Version_HARDIP = "Rhea_A0_Vector_List_HardIP#88_EXT_140922"
'''''''    Else
'''''''        Excel_Version = "Cayman_A0_PatternList_scgh_150913_22281924121762"
'''''''    End If
    
    If gL_ProductionTemp = "" Then 'come from global variant
        gL_ProductionTemp = "Null"
    End If
    
    If gS_SPI_Version = "" Then 'come from global variant
        gS_SPI_Version = "Null"
    End If
    
    'TheExec.Datalog.WriteComment "*******************program information*******************"
    TheExec.Datalog.WriteComment "********************************"
    TheExec.Datalog.WriteComment "*print: Program information    *"
    TheExec.Datalog.WriteComment "********************************"
    
    'job printing
    TheExec.Datalog.WriteComment "print: Current job name, " & UCase(currentJobName)
    'pattern version print
    TheExec.Datalog.WriteComment "print: Pattern list version, " & PatVersion
    'execl version
    If separate = True Then
        TheExec.Datalog.WriteComment "print: SocScan, " & Excel_Version_SOC_SCAN
        TheExec.Datalog.WriteComment "print: SocMbist, " & Excel_Version_SOC_MBIST
        TheExec.Datalog.WriteComment "print: CpuScan, " & Excel_Version_CPU_SCAN
        TheExec.Datalog.WriteComment "print: CpuMbist, " & Excel_Version_CPU_MBIST
        TheExec.Datalog.WriteComment "print: CpuPCM, " & Excel_Version_CPU_PCM
        TheExec.Datalog.WriteComment "print: HardIP, " & Excel_Version_HARDIP
    Else
        TheExec.Datalog.WriteComment "print: All Excel, " & Excel_Version
    End If
    TheExec.Datalog.WriteComment "print: Test plan version, " & TestPlanVersion
    TheExec.Datalog.WriteComment "print: Test temparature, " & gL_ProductionTemp & " degC"
    TheExec.Datalog.WriteComment "print: SPIROM version, " & gS_SPI_Version
    If LCase(TheExec.CurrentJob) Like "*char*" Then
    Dim gS_Char_Version As String
    Set wb1 = Application.ActiveWorkbook
    Set ws_pat1 = wb1.Sheets("Flow_Char")
    gS_Char_Version = ws_pat1.Cells(7, 8).Value
        TheExec.Datalog.WriteComment "print: Char plan version, " & gS_Char_Version
    End If
    
    'TheExec.Datalog.WriteComment "*******************program information*******************"
End Function

'*****************************************
'******            Read/Write EPPROM******
'*****************************************
Public Function Write_DIB_EEPROM(Optional DIB_SerialNumber As String) As Long
    On Error GoTo errHandler
    Dim CurrJob As String
    Dim config(2) As Long
    config(0) = 32768
    config(1) = 0
    config(2) = 0
    
    TheHdw.DIB.PIBEEPROM.program (config)
'''    CurrJob = TheExec.CurrentJob    'CP/FT judge
'''    change to LCase to prevent ft1 FT1 issue

    If TheExec.Flow.EnableWord("Write_EEPROM_DIBID") = True Then
'''     If (CurrJob Like "cp*") Then
        If LCase(TheExec.CurrentJob) Like "cp*" Then
            If RegKeyRead("PROBECARD_ID") <> "" Then
                DIB_SerialNumber = RegKeyRead("PROBECARD_ID")
            Else
                DIB_SerialNumber = ""
            End If
'''     ElseIf (CurrJob Like "ft*") Then
        ElseIf LCase(TheExec.CurrentJob) Like "ft*" Then
            If RegKeyRead("LOADBOARD_ID") <> "" Then
                DIB_SerialNumber = RegKeyRead("LOADBOARD_ID")
            Else
                DIB_SerialNumber = ""
            End If
        End If
        
        TheHdw.DIB.PIBEEPROM.Record("DIB_SerialNum") = DIB_SerialNumber
        TheHdw.DIB.PIBEEPROM.Record.WriteToHW   'write to HW
        TheExec.Datalog.WriteComment ("print: Write DIB ID " & DIB_SerialNumber)
    End If  'end enable wd block
    'only fuse once
    TheExec.Flow.EnableWord("Write_EEPROM_DIBID") = False
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

'Public Function Write_SPIROM_Check() As Boolean
'
'    Write_SPIROM_Check = False
'
'    Dim Site As Variant
'
'    For Each Site In TheExec.Sites
'        If (write_SPIROM_CheckSum And (2 ^ Site)) = 0 Then
'            Write_SPIROM_Check = True
'            write_SPIROM_CheckSum = (write_SPIROM_CheckSum Or (2 ^ Site))
'        End If
'    Next Site
'
'End Function
Public Function Read_DIB_EEPROM() As Long
    On Error GoTo errHandler
    
    Dim rec() As IDIB_EEPROM_RecordObj
    Dim i As Integer
    Dim DIB_ID As String
    
    If TheHdw.DIB.PIBEEPROM.IsProgrammed Then
        'Debug.Print "The PIB EEPROM is programmed"
        rec = TheHdw.DIB.PIBEEPROM.Record.List
    
'''        For i = 0 To UBound(rec)
'''            Debug.Print "Record " + rec(i).ID + " = " + rec(i).Value
'''        Next i

        DIB_ID = rec(0).Value
        TheExec.Datalog.WriteComment ("print: Read DIB ID is " & DIB_ID)
    Else
        'Debug.Print "The PIB EEPROM is not programmed"
        TheExec.Datalog.WriteComment ("print: Read DIB ID is not programmed, read fail")
    End If
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ReadProberTemp(Temp_Hilimit As Double, Temp_Lolimit As Double)

Dim s As Variant
Dim Prober_Temp As Double
Dim Prober_Temp_str As String
Dim Count As Integer
Dim i As Integer, j As Integer
Dim Num_Str(9) As String
Dim Temp_Str_status As Integer
On Error GoTo errHandler
        TheExec.Datalog.WriteComment ""
        'TheExec.Datalog.WriteComment "/********* Read Prober Temperature *********/"
        TheExec.Datalog.WriteComment "********************************"
        TheExec.Datalog.WriteComment "*print: Read Prober Temperature*"
        TheExec.Datalog.WriteComment "********************************"
        
        If IsNumeric(RegKeyRead("Prober_Temp")) = True Then
            Prober_Temp_str = RegKeyRead("Prober_Temp")

        Else
            If RegKeyRead("Prober_Temp") = "" Then
                TheExec.Datalog.WriteComment "Registry is empty"
                Prober_Temp_str = "00000"
            Else
                TheExec.Datalog.WriteComment "Registry is not empty nor number."
                Prober_Temp_str = "99999"
            End If
        End If
        
        If Mid(Prober_Temp_str, 1, 1) Like "+" Then
                Prober_Temp = CDbl(Mid(Prober_Temp_str, 2, 5))
        Else
                Prober_Temp = CDbl(Prober_Temp_str)
        End If
        
        gL_ProductionTemp = Prober_Temp 'upload to global variant
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            Prober_Temp = 25
            gL_ProductionTemp = Prober_Temp
        End If
        
        TheExec.Datalog.WriteComment "Prober_Temp(Registry) : " & RegKeyRead("Prober_Temp")
        TheExec.Flow.TestLimit Prober_Temp, lowVal:=Temp_Lolimit, hiVal:=Temp_Hilimit, Tname:="Prober_Temp_" & CStr(Mid(TheExec.DataManager.instanceName, 16, 18))
        'TheExec.Datalog.WriteComment "/*********************************************/"
       
    Exit Function
errHandler:
        TheExec.Datalog.WriteComment "Read Prober Temp VBT function is error "
        TheExec.Datalog.WriteComment "Registry String :" & RegKeyRead("Prober_Temp")
        TheExec.Datalog.WriteComment ("Error #: " & Str(err.number) & " " & err.Description)
        If AbortTest Then Exit Function Else Resume Next
End Function


Public Function SetupInitialCondition(DisableConnectPins As PinList, DrivePins As PinList, Optional DriveVolt As Double = 0.1) As Long
    If (DisableConnectPins <> "") Then
        TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    End If
    If (DrivePins <> "") Then
        TheHdw.Digital.Pins(DrivePins).Disconnect
        TheHdw.Wait (100 * us)
        
        With TheHdw.PPMU.Pins(DrivePins)
            .Gate = tlOff
            .ForceV DriveVolt, 0.002
            .Connect
            TheHdw.Wait (100 * us)
            .Gate = tlOn
        End With
    End If
End Function


Public Function Set_PPMU_Clamp(Pin_GP1 As PinList, Pin_GP1_Vch As Double, _
                               Pin_GP2 As PinList, Pin_GP2_Vch As Double, _
                               Pin_GP3 As PinList, Pin_GP3_Vch As Double, _
                               Pin_GP4 As PinList, Pin_GP4_Vch As Double, _
                               Pin_GP5 As PinList, Pin_GP5_Vch As Double, _
                               Pin_GP6 As PinList, Pin_GP6_Vch As Double, _
                               Pin_GP7 As PinList, Pin_GP7_Vch As Double, _
                               Pin_GP8 As PinList, Pin_GP8_Vch As Double, _
                               Pin_GP9 As PinList, Pin_GP9_Vcl As Double, _
                               Pin_GP10 As PinList, Pin_GP10_Vcl As Double)


On Error GoTo errHandler


    
    If Pin_GP1.Value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP1).ClampVHi = Pin_GP1_Vch
    End If
    
    If Pin_GP2.Value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP2).ClampVHi = Pin_GP2_Vch
    End If
    
    If Pin_GP3.Value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP3).ClampVHi = Pin_GP3_Vch
    End If

    If Pin_GP4.Value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP4).ClampVHi = Pin_GP4_Vch
    End If
 
    If Pin_GP5.Value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP5).ClampVHi = Pin_GP5_Vch
    End If
 
    If Pin_GP6.Value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP6).ClampVHi = Pin_GP6_Vch
    End If
 
    If Pin_GP7.Value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP7).ClampVHi = Pin_GP7_Vch
    End If
 
    If Pin_GP8.Value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP8).ClampVHi = Pin_GP8_Vch
    End If
 
    If Pin_GP9.Value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP9).ClampVLo = Pin_GP9_Vcl
    End If
    If Pin_GP10.Value = "" Then
    Else
        TheHdw.PPMU.Pins(Pin_GP10).ClampVLo = Pin_GP10_Vcl
    End If
 
    Exit Function
 
errHandler:

                If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Read_Package_ID()

Dim site As Variant

On Error GoTo errHandler

        TheExec.Datalog.WriteComment "********************************"
        
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site:" & site & " Device ID : " & RegKeyRead("ManualTestDeviceID"))
        Next site
        
        TheExec.Datalog.WriteComment "********************************"
        
        Dim testName As String
        testName = "Device ID"
        Dim ResultVal_DeviceID As Double
        ResultVal_DeviceID = CDbl(RegKeyRead("ManualTestDeviceID"))
        For Each site In TheExec.sites
            
            TheExec.Flow.TestLimit resultVal:=ResultVal_DeviceID, Tname:=testName, ForceResults:=tlForceNone
        Next site
        
       
    Exit Function
errHandler:
        TheExec.Datalog.WriteComment "Read Package ID error "
        If AbortTest Then Exit Function Else Resume Next
End Function


Public Function FreeRunClk_Disable_MultiPort(PortName As String) As Long '''update for multi nWire 20170718
    Dim site As Variant
    Dim i As Integer
    Dim A_PortName() As String
    
    Call Disable_FRC("", False)
''    theexec.Datalog.WriteComment "******************  Disable freerunning clock ****************"
''
''    A_PortName = Split(PortName, ",")
''
''    For i = 0 To UBound(A_PortName)
''        ' Disable the nWire engine.
''        'If TheExec.Flow.EnableWord("XI0_nWire") = True Then
''            TheHdw.Protocol.ports(A_PortName(i)).Halt
''            TheHdw.Protocol.ports(A_PortName(i)).Enabled = False
''        'End If
'''        If A_PortName(i) Like "Clock_Port" Then
'''            TheHdw.Digital.Pins("Xi0_PA").InitState = chInitLo
'''        ElseIf PortName Like "RTCLK_Port" Then
'''            TheHdw.Digital.Pins("RT_CLK32768_PA").InitState = chInitLo
'''        End If
'''        TheHdw.Wait 0.001
'''        TheHdw.Digital.Pins(A_PortName(i)).InitState = chInitoff
''
''        'TheHdw.DIB.SupportBoardClock.Stop
''    Next i
''
''    ''upload to global constant
''    FreeRunFreq_debug = 0
''    clock_Vih_debug = 0
''    clock_Vil_debug = 0

End Function


Public Function CheckFlag(Optional B_PrintFlag As Boolean = False)

    Dim F_Sa_Flag As String
    Dim F_SaChain_Flag As String
    Dim F_TD_Flag As String
    Dim F_Bist_Flag As String
    Dim F_Bist_MC_Flag As String
    Dim F_SaHV_Flag As String
    Dim F_SaLV_Flag As String
    Dim F_SaChainHV_Flag As String
    Dim F_SaChainLV_Flag As String
    Dim F_Pass As String
    Dim F_Fail As String
    Dim site As Variant
    Dim DLY_PassCount As Long
    Dim DLY_Top1_PassCount As Long: DLY_Top1_PassCount = 16
    Dim DLY_Top2_PassCount As Long: DLY_Top2_PassCount = 16
    Dim PLY_PassCount As Long
    Dim PLY_Top1_PassCount As Long: PLY_Top1_PassCount = 16
    Dim PLY_Top2_PassCount As Long: PLY_Top2_PassCount = 16
    Dim S_FlagName As String
    Dim S_FlagState As String
    Dim S_Sa_FlagState As String
    Dim S_SaChain_FlagState As String
    Dim S_TD_FlagState As String
    Dim S_Bist_FlagState As String
    Dim S_Bist_MC_FlagState As String
    Dim S_SaHV_FlagState As String
    Dim S_SaLV_FlagState As String
    Dim S_SaChainHV_FlagState As String
    Dim S_SaChainLV_FlagState As String
    Dim B_FirstSite As Boolean
    Dim i As Integer
    Dim A_FlagName() As String
    Dim A_FlagState() As String
    
    B_FirstSite = True
    
    For Each site In TheExec.sites
        S_FlagName = ""
        S_FlagState = ""
        DLY_Top1_PassCount = 16
        DLY_Top2_PassCount = 16
        PLY_Top1_PassCount = 16
        PLY_Top2_PassCount = 16
        
        For i = 0 To 31
            'Right(CStr(i + 100), 2) -> to get 2 character while convert long to string
            F_Sa_Flag = "F_cpusa_B" & Right(CStr(i + 100), 2) & "_MHLV"
            F_SaChain_Flag = "F_cpusachain_B" & Right(CStr(i + 100), 2) & "_MHLV"
            F_TD_Flag = "F_CpuTD_B" & Right(CStr(i + 100), 2) & "_NV"
            F_Bist_Flag = "F_cpubist_B" & Right(CStr(i + 100), 2) & "_MHLV"
            F_Bist_MC_Flag = "F_cpubist_B" & Right(CStr(i + 100), 2) & "_NV"
            
            F_SaHV_Flag = "F_cpusa_B" & Right(CStr(i + 100), 2) & "_HV"
            F_SaLV_Flag = "F_cpusa_B" & Right(CStr(i + 100), 2) & "_LV"
            F_SaChainHV_Flag = "F_cpusachain_B" & Right(CStr(i + 100), 2) & "_HV"
            F_SaChainLV_Flag = "F_cpusachain_B" & Right(CStr(i + 100), 2) & "_LV"
            
            ' set SA MHLV flag
            If TheExec.sites.Item(site).FlagState(F_SaHV_Flag) = logicFalse Or TheExec.sites.Item(site).FlagState(F_SaLV_Flag) = logicFalse Then
                TheExec.sites.Item(site).FlagState(F_Sa_Flag) = logicFalse
            ElseIf TheExec.sites.Item(site).FlagState(F_SaHV_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_SaLV_Flag) = logicTrue Then
                TheExec.sites.Item(site).FlagState(F_Sa_Flag) = logicTrue
            Else
                TheExec.sites.Item(site).FlagState(F_Sa_Flag) = logicClear
            End If
            
            ' set SA Chain MHLV flag
            If TheExec.sites.Item(site).FlagState(F_SaChainHV_Flag) = logicFalse Or TheExec.sites.Item(site).FlagState(F_SaChainLV_Flag) = logicFalse Then
                TheExec.sites.Item(site).FlagState(F_SaChain_Flag) = logicFalse
            ElseIf TheExec.sites.Item(site).FlagState(F_SaChainHV_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_SaChainLV_Flag) = logicTrue Then
                TheExec.sites.Item(site).FlagState(F_SaChain_Flag) = logicTrue
            Else
                TheExec.sites.Item(site).FlagState(F_SaChain_Flag) = logicClear
            End If
            
            If i < 16 Then
                ' if any of SA, SAChain, Bist_MHLV fail on this Top1 cpu core, then this cpu core fail
                If TheExec.sites.Item(site).FlagState(F_Sa_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_SaChain_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_Bist_Flag) = logicTrue Then
                    DLY_Top1_PassCount = DLY_Top1_PassCount - 1
                End If
                ' if any of SA, SAChain, TD, Bist, Bist_MC fail on this Top1 cpu core, then this cpu core fail
                If TheExec.sites.Item(site).FlagState(F_Sa_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_SaChain_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_TD_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_Bist_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_Bist_MC_Flag) = logicTrue Then
                    PLY_Top1_PassCount = PLY_Top1_PassCount - 1
                End If
            Else
                ' if any of SA, SAChain, Bist_MHLV fail on this Top2 cpu core, then this cpu core fail
                If TheExec.sites.Item(site).FlagState(F_Sa_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_SaChain_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_Bist_Flag) = logicTrue Then
                    DLY_Top2_PassCount = DLY_Top2_PassCount - 1
                End If
                ' if any of SA, SAChain, TD, Bist, Bist_MC fail on this Top2 cpu core, then this cpu core fail
                If TheExec.sites.Item(site).FlagState(F_Sa_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_SaChain_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_TD_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_Bist_Flag) = logicTrue Or TheExec.sites.Item(site).FlagState(F_Bist_MC_Flag) = logicTrue Then
                    PLY_Top2_PassCount = PLY_Top2_PassCount - 1
                End If
            End If
            
            ' collect flag state
            If B_PrintFlag = True Then
                If TheExec.sites.Item(site).FlagState(F_SaHV_Flag) = logicTrue Then
                    S_SaHV_FlagState = "T"
                ElseIf TheExec.sites.Item(site).FlagState(F_SaHV_Flag) = logicFalse Then
                    S_SaHV_FlagState = "F"
                Else
                    S_SaHV_FlagState = "C"
                End If
                If TheExec.sites.Item(site).FlagState(F_SaLV_Flag) = logicTrue Then
                    S_SaLV_FlagState = "T"
                ElseIf TheExec.sites.Item(site).FlagState(F_SaLV_Flag) = logicFalse Then
                    S_SaLV_FlagState = "F"
                Else
                    S_SaLV_FlagState = "C"
                End If
                If TheExec.sites.Item(site).FlagState(F_Sa_Flag) = logicTrue Then
                    S_Sa_FlagState = "T"
                ElseIf TheExec.sites.Item(site).FlagState(F_Sa_Flag) = logicFalse Then
                    S_Sa_FlagState = "F"
                Else
                    S_Sa_FlagState = "C"
                End If
                If TheExec.sites.Item(site).FlagState(F_SaChainHV_Flag) = logicTrue Then
                    S_SaChainHV_FlagState = "T"
                ElseIf TheExec.sites.Item(site).FlagState(F_SaChainHV_Flag) = logicFalse Then
                    S_SaChainHV_FlagState = "F"
                Else
                    S_SaChainHV_FlagState = "C"
                End If
                If TheExec.sites.Item(site).FlagState(F_SaChainLV_Flag) = logicTrue Then
                    S_SaChainLV_FlagState = "T"
                ElseIf TheExec.sites.Item(site).FlagState(F_SaChainLV_Flag) = logicFalse Then
                    S_SaChainLV_FlagState = "F"
                Else
                    S_SaChainLV_FlagState = "C"
                End If
                If TheExec.sites.Item(site).FlagState(F_SaChain_Flag) = logicTrue Then
                    S_SaChain_FlagState = "T"
                ElseIf TheExec.sites.Item(site).FlagState(F_SaChain_Flag) = logicFalse Then
                    S_SaChain_FlagState = "F"
                Else
                    S_SaChain_FlagState = "C"
                End If
                If TheExec.sites.Item(site).FlagState(F_TD_Flag) = logicTrue Then
                    S_TD_FlagState = "T"
                ElseIf TheExec.sites.Item(site).FlagState(F_TD_Flag) = logicFalse Then
                    S_TD_FlagState = "F"
                Else
                    S_TD_FlagState = "C"
                End If
                If TheExec.sites.Item(site).FlagState(F_Bist_Flag) = logicTrue Then
                    S_Bist_FlagState = "T"
                ElseIf TheExec.sites.Item(site).FlagState(F_Bist_Flag) = logicFalse Then
                    S_Bist_FlagState = "F"
                Else
                    S_Bist_FlagState = "C"
                End If
                If TheExec.sites.Item(site).FlagState(F_Bist_MC_Flag) = logicTrue Then
                    S_Bist_MC_FlagState = "T"
                ElseIf TheExec.sites.Item(site).FlagState(F_Bist_MC_Flag) = logicFalse Then
                    S_Bist_MC_FlagState = "F"
                Else
                    S_Bist_MC_FlagState = "C"
                End If
                
                S_FlagName = S_FlagName & "," & F_SaHV_Flag & "," & F_SaLV_Flag & "," & F_Sa_Flag & "," & F_SaChainHV_Flag & "," & F_SaChainLV_Flag & "," & F_SaChain_Flag & "," & F_TD_Flag & "," & F_Bist_Flag & "," & F_Bist_MC_Flag
                S_FlagState = S_FlagState & "," & S_SaHV_FlagState & "," & S_SaLV_FlagState & "," & S_Sa_FlagState & "," & S_SaChainHV_FlagState & "," & S_SaChainLV_FlagState & "," & S_SaChain_FlagState & "," & S_TD_FlagState & "," & S_Bist_FlagState & "," & S_Bist_MC_FlagState
            End If
            
        Next i
        
        
        If DLY_Top1_PassCount < DLY_Top2_PassCount Then
            DLY_PassCount = DLY_Top1_PassCount
        Else
            DLY_PassCount = DLY_Top2_PassCount
        End If
        
        If PLY_Top1_PassCount < PLY_Top2_PassCount Then
            PLY_PassCount = PLY_Top1_PassCount
        Else
            PLY_PassCount = PLY_Top2_PassCount
        End If
        
        
        F_Pass = "P_DLY_Cpu_Core_" & DLY_PassCount & "_pass"
        TheExec.sites.Item(site).FlagState(F_Pass) = logicTrue
        
        F_Pass = "P_PLY_Cpu_Core_" & PLY_PassCount & "_pass"
        TheExec.sites.Item(site).FlagState(F_Pass) = logicTrue
        
        F_Pass = "P_PLY_Top1_Cpu_Core_" & PLY_Top1_PassCount & "_pass"
        TheExec.sites.Item(site).FlagState(F_Pass) = logicTrue
        
        F_Pass = "P_PLY_Top2_Cpu_Core_" & PLY_Top2_PassCount & "_pass"
        TheExec.sites.Item(site).FlagState(F_Pass) = logicTrue
        
        If B_PrintFlag = True Then
            If S_FlagName <> "" Then S_FlagName = Right(S_FlagName, Len(S_FlagName) - 1)
            If S_FlagState <> "" Then S_FlagState = Right(S_FlagState, Len(S_FlagState) - 1)
            
            A_FlagName = Split(S_FlagName, ",")
            A_FlagState = Split(S_FlagState, ",")
            
'            If B_FirstSite = True Then
                For i = 0 To UBound(A_FlagName)
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & A_FlagName(i) & ", " & A_FlagState(i)
'                    TheExec.Datalog.WriteComment "FlagState_Site(" & Site & ")    " & S_FlagState
                Next i
'                B_FirstSite = False
'            Else
'                TheExec.Datalog.WriteComment "FlagState_Site(" & Site & ")    " & S_FlagState
'            End If


        If TheExec.sites.Item(site).FlagState("F_socmbist_bist_hard_defect") = logicTrue Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_hard_defect" & ", T"
                ElseIf TheExec.sites.Item(site).FlagState("F_socmbist_bist_hard_defect") = logicFalse Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_hard_defect" & ", F"
                Else
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_hard_defect" & ", C"
                End If
                
                If TheExec.sites.Item(site).FlagState("F_socmbist_bist_bitcell") = logicTrue Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_bitcell" & ", T"
                ElseIf TheExec.sites.Item(site).FlagState("F_socmbist_bist_bitcell") = logicFalse Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_bitcell" & ", F"
                Else
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_socmbist_bist_bitcell" & ", C"
                End If
        
                If TheExec.sites.Item(site).FlagState("F_soc_fail") = logicTrue Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_soc_fail" & ", T"
                ElseIf TheExec.sites.Item(site).FlagState("F_soc_fail") = logicFalse Then
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_soc_fail" & ", F"
                Else
                    TheExec.Datalog.WriteComment "Site(" & site & "), " & "F_soc_fail" & ", C"
                End If
                
        End If
        
        
    Next site
    
End Function


Public Function Alarm_binout()


    Dim site As Variant

    'bin out alarm
    For Each site In TheExec.sites

       If TheHdw.Alarms.Check = True Then

                'TheExec.Sites.Item(Site).SortNumber = 908
                'TheExec.Sites.Item(Site).BinNumber = 6
                TheExec.sites.Item(site).result = tlResultFail

       End If

    Next site


    Exit Function
    
errHandler:

      If AbortTest Then Exit Function Else Resume Next

End Function
Public Function Set_DCVS_alarm(Pin As PinList, AlarmTime As Double, Pin_GP2 As PinList, AlarmTime_GP2 As Double, Optional b_AlarmForceBin As Boolean = False)
        'add boolean to set force_bin or force_fail

On Error GoTo errHandler

    If Pin.Value = "" And Pin_GP2.Value = "" Then
        TheExec.Datalog.WriteComment "*******************************************************************"
        TheExec.Datalog.WriteComment "Error on Set_DCVS_alarm, please fill in pin name to set alarm time."
        TheExec.Datalog.WriteComment "*******************************************************************"
    End If

    With TheHdw.DCVS.Pins(Pin)
        .Gate = False
        TheHdw.Wait 1 * ms '20170209 Add to wait gate off to avoid set ifold time out error
        .mode = tlDCVSModeVoltage
        .Voltage.Main = 0
        
        If b_AlarmForceBin = False Then
            .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail
        Else
            .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin
        End If
        If CStr(AlarmTime) = "" Or AlarmTime < 0.0001 Then
            TheExec.Datalog.WriteComment "************************************************************"
            TheExec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time."
            TheExec.Datalog.WriteComment "************************************************************"
        End If
        '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime
        .CurrentLimit.Source.FoldLimit.TimeOut = AlarmTime
    End With
    
    If Pin_GP2.Value = "" Then
    Else
        With TheHdw.DCVS.Pins(Pin_GP2)
            .Gate = False
            TheHdw.Wait 1 * ms '20170209 Add to wait gate off to avoid set ifold time out error
            .mode = tlDCVSModeVoltage
            .Voltage.Main = 0
            
            If b_AlarmForceBin = False Then
                .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail
            Else
                .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin
            End If
            If CStr(AlarmTime_GP2) = "" Or AlarmTime_GP2 < 0.0001 Then
                TheExec.Datalog.WriteComment "*******************************************************************"
                TheExec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time on GP2."
                TheExec.Datalog.WriteComment "*******************************************************************"
            End If
            '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime
            .CurrentLimit.Source.FoldLimit.TimeOut = AlarmTime_GP2
        End With
    End If
    
    
    Dim CurrentChans As String
     CurrentChans = TheExec.CurrentChanMap 'obtain FT or CP channel map information
    If CurrentChans Like "*FT*" Then
       'AllPowerPinlist = "All_Power_FT"
       'Utility_list = "FT_Utility"
        'increase the foldlimit time out for Vdd_cpu_sram
'        With thehdw.DCVS.pins("vdd_sram")
'           .CurrentLimit.Source.FoldLimit.TimeOut = 0.02
'
'        End With
    Else
     'increase the foldlimit time out for Vdd_cpu_sram
'        With thehdw.DCVS.pins("vdd_cpu_sram")
'           .CurrentLimit.Source.FoldLimit.TimeOut = 0.02
'
'        End With
      ' AllPowerPinlist = "All_Power_CP"
       'Utility_list = "CP_Utility"
    End If

 
    Exit Function
 
errHandler:

                If AbortTest Then Exit Function Else Resume Next
End Function
Public Function KA_start() As Long
    TheHdw.Digital.Patgen.KeepAlive.Flag = cpuB
    TheHdw.Digital.Patgen.KeepAlive.Enable = True
End Function

Public Function KA_end() As Long
    TheHdw.Digital.Patgen.Halt
    Call TheHdw.Digital.Patgen.Continue(0, cpuB)
    TheHdw.Digital.Patgen.KeepAlive.Enable = False
End Function
Public Function Disable_compare(DisableCompare_Pin As PinList)

    TheHdw.Digital.Pins(DisableCompare_Pin).DisableCompare = True
    
    TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Datalog.WriteComment "*Disable Compare Pin:" & DisableCompare_Pin
    TheExec.Datalog.WriteComment "*************************************************"
    
End Function

Public Function Enble_compare(EnableCompare_Pin As PinList)

    TheHdw.Digital.Pins(EnableCompare_Pin).DisableCompare = False

    TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Datalog.WriteComment "*Enable Compare Pin:" & EnableCompare_Pin
    TheExec.Datalog.WriteComment "*************************************************"
    
End Function
Public Function Plot_Profile_on_disk(PinName As PinList, SampleRate As Double, SampleSize As Long, Optional CapSignalName As String = "Capture_signal")
'Plot profiles

    Dim DSPW As New DSPWave
    Dim Label As String
    Dim site As Variant
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim ForCount As Long
    Dim ForCount2 As Long
    Dim p As Variant
    Dim FolderName As String
    Dim FileName As String
    Dim FirstActiveSite As Long
    Dim RowString As String
    Dim t1, t2 As Double
    Dim DebugLog As Boolean
    Dim dataArray() As String
    FirstActiveSite = -1
    DebugLog = True
    
    On Error GoTo errHandler

    Do While TheHdw.DCVS.Pins(PinName).Capture.IsRunning = True
    Loop

    ' Get the captured samples from the instrument
    Call TheExec.DataManager.DecomposePinList(PinName, Pin_Ary(), Pin_Cnt)

    ReDim dataArray(SampleSize, Pin_Cnt + 1)

    FolderName = "D:\"
    t1 = Timer
    
    For ForCount = 1 To SampleSize
        If ForCount > 1000 And ForCount > SampleSize - 100 Then Exit For 'avoid less capture
        dataArray(ForCount + 1, 1) = (ForCount - 1) / SampleRate * 1000
    Next ForCount

    t2 = Timer
    If DebugLog Then TheExec.Datalog.WriteComment ("    Print First column time : " & (t2 - t1) * 1000 & " ms.")

    ForCount = 0
    For Each p In Pin_Ary
        If TheExec.DataManager.ChannelType(p) <> "N/C" Then
            DSPW = TheHdw.DCVS.Pins(p).Capture.Signals(CapSignalName).DSPWave
            For Each site In TheExec.sites
                If TheHdw.DCVS.Pins(p).Meter.mode = tlDCVSMeterCurrent Then
                    Label = "Current Profile for Site: " & site & " " & p
                    FileName = FolderName & site & "_Pin_" & p & ".txt"
                Else
                    Label = "Voltage Profile for Site: " & site & " " & p
                    FileName = FolderName & site & "_Pin_" & p & ".txt"
                End If
                
                If False Then DSPW.plot Label   'for pliot
                If False Then DSPW.FileExport FileName, File_txt  'for file export
                If True Then 'for Save All pin Data into output file
                    If FirstActiveSite = -1 Then
                        FirstActiveSite = site
                    End If
                    If site = FirstActiveSite Then
                        t1 = Timer
                        dataArray(1, ForCount + 2) = p
                        For ForCount2 = 1 To SampleSize
                            If ForCount2 > 1000 And ForCount2 > SampleSize - 100 Then Exit For 'avoid less capture
                            dataArray(ForCount2 + 1, ForCount + 2) = DSPW.Element(ForCount2)
                        Next ForCount2
                        t2 = Timer
                        If DebugLog Then TheExec.Datalog.WriteComment ("    Plot Profile Pin: " & p & " is done. Execute Time : " & (t2 - t1) * 1000 & " ms.")
                        ForCount = ForCount + 1
                    End If

                End If
            Next site
        End If
    Next p
    
    t1 = Timer
    FileName = FolderName & "VIProfile_site" & FirstActiveSite & "_" & t1 \ 1 & ".txt"
    Open FileName For Output As #1
    For ForCount = 1 To SampleSize - 100
        RowString = ""
        For ForCount2 = 1 To Pin_Cnt + 1
            RowString = RowString & dataArray(ForCount, ForCount2) & " "
        Next ForCount2
        Print #1, RowString
    Next ForCount
    t2 = Timer
    If DebugLog Then TheExec.Datalog.WriteComment ("    Plot Profile Pin: " & p & " is done. Execute Time : " & (t2 - t1) * 1000 & " ms.")
    
    Close #1
    
    TheExec.Datalog.WriteComment ("    Plot Profile All Pins are done.")
    
    Exit Function
errHandler:
    ErrorDescription ("Plot_Profile_NEW")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function SiteResultCheck()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SiteResultCheck"

    Dim site As Variant
    Dim b_SitePass As New SiteLong

    For Each site In TheExec.sites
        If TheExec.sites.Item(site).result = tlResultPass Then
            b_SitePass = 1
        ElseIf TheExec.sites.Item(site).result = tlResultFail Then
            b_SitePass = 0
        Else
            b_SitePass = 1
        End If
    Next site

    TheExec.Flow.TestLimit b_SitePass, 1, 1

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function FreeRunclk_Enable(PortName As String)

    Dim site As Variant
    Dim i As Long
    Dim PLLLockChecked As New SiteLong
    Dim measf As New PinListData
    
    Dim NotLocked As Boolean
    Dim XI0_REFCLK As String
    Dim PortMode As String
    Dim PLL_Lock As New SiteLong
    Dim port_level_name As String
    Dim port_level_value As Double
    Dim FreeRunFreq As Double
    On Error GoTo errHandler
    
    'Call Disable_FRC(PortName, False)
    
    If TheExec.DataManager.instanceName <> "" Then
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    
    Call Enable_FRC(PortName, False)
    TheHdw.Wait 0.001
  
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function PowerDown_Parallel(PowerPinList_DCVS As String, PowerPinList_DCVI As String, DisconnectPinList As String, Optional WaitConnectTime As Double = 0.001, Optional DebugFlag As Boolean = True) ', _
                Optional DriveLowPinList As PinList, Optional ClockPat As Pattern, Optional RTCLK_Relay As PinList, Optional XI0_Relay As PinList)
    Dim CurrentChans As String
    Dim Pins() As String, PinCnt As Long
    Dim powerPin As Variant
    Dim PowerName As String
    Dim TempString As String
    Dim Vmain As Double
    Dim Irange As Double
    Dim step As Integer
    Dim FallTime As Double
    Dim PowerSequence As Double
    Dim i As Long, j As Long
    Dim site As Variant
    
    Dim RTCLK_Relay As New PinList
    Dim XI0_Relay As New PinList
    
    Dim PowerSequencePin() As String
    Dim seqnum As Long
    Dim TempMaxSequence As Long:: TempMaxSequence = 0
     
    Dim nwire_seq() As Long
    Dim nwire_port() As String, nwire_port_plist As New PinList
    Dim nwire_pin() As String
    Dim nWire_port_ary() As String
    Dim nWire_port_count As Long
    Dim nwp As Variant, nw_i As Long
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim pins_dcvs() As String, PinCnt_dcvs As Long, pins_dcvi() As String, PinCnt_dcvi As Long

    Dim pin_name As String
    Dim SlotType As String

    Dim IO_H_seq_nu() As Long, IO_L_seq_nu() As Long, IO_HZ_seq_nu() As Long
    Dim IO_H_seq_pin() As String, IO_L_seq_pin() As String, IO_HZ_seq_pin() As String
    Dim IO_H_seq_pin_total() As String, IO_L_seq_pin_total() As String, IO_HZ_seq_pin_total() As String
    Dim IO_H_list() As String, IO_L_list() As String, IO_HZ_list() As String
    Dim IO_H_count As Long, IO_L_count As Long, IO_HZ_count As Long
    Dim IO_H_nu As Long, IO_L_nu As Long, IO_HZ_nu As Long

    On Error GoTo errHandler
                        
    Call Print_Header("Power down sequence")
    
    If power_up_en <> True Then
        '''''''''''''''''Support multiple nWire port 20170503'''''''''''''
        nWire_port_ary = Split(nWire_Ports_GLB, ",")
        nWire_port_count = 0
        ReDim nwire_port(UBound(nWire_port_ary))
        ReDim nwire_seq(UBound(nWire_port_ary))
        ReDim nwire_pin(UBound(nWire_port_ary))
        For Each nwp In nWire_port_ary
            Get_nWire_Name CStr(nwp), port_pa, ac_spec_pa, pin_pa, global_spec_pa
            nwire_port(nWire_port_count) = port_pa
            nwire_seq(nWire_port_count) = TheExec.specs.Globals(global_spec_pa).ContextValue
            nwire_pin(nWire_port_count) = pin_pa
            nWire_port_count = nWire_port_count + 1
        Next nwp
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If

    TheExec.Datalog.WriteComment vbCrLf & "print: Power down start, Power pins: " & PowerPinList_DCVS & "," & PowerPinList_DCVI
    TheHdw.Digital.Pins(DisconnectPinList).Disconnect
    TheExec.Datalog.WriteComment "print: Power down digital disconnect, Digital pins: " & DisconnectPinList
    TheExec.Datalog.WriteComment RepeatChr("*", 120)
    
    Dim Power_up_down_flag As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If power_dcvs_exit = True Then
        TheExec.DataManager.DecomposePinList PowerPinList_DCVS, pins_dcvs(), PinCnt_dcvs
        If TheExec.specs.Globals.Contains(pins_dcvs(0) & "_PowerDownSequence_GLB") = True Then Power_up_down_flag = True
    End If
    If power_dcvi_exit = True Then
        TheExec.DataManager.DecomposePinList PowerPinList_DCVI, pins_dcvi(), PinCnt_dcvi
        If TheExec.specs.Globals.Contains(pins_dcvi(0) & "_PowerDownSequence_GLB") = True Then Power_up_down_flag = True
    End If
    
    If io_h_pins <> "" Then
        TheExec.DataManager.DecomposePinList io_h_pins, IO_H_seq_pin(), IO_H_count
    End If
    If io_l_pins <> "" Then
        TheExec.DataManager.DecomposePinList io_l_pins, IO_L_seq_pin(), IO_L_count
    End If
    If io_hz_pins <> "" Then
        TheExec.DataManager.DecomposePinList io_hz_pins, IO_HZ_seq_pin(), IO_HZ_count
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    PinCnt = PinCnt_dcvs + PinCnt_dcvi
    ReDim Pins(PinCnt - 1)
    
    For i = 0 To PinCnt - 1
        If i < PinCnt_dcvs And PinCnt_dcvs > 0 Then
            Pins(i) = pins_dcvs(i)
        Else
            Pins(i) = pins_dcvi(i - PinCnt_dcvs)
        End If
    Next i
    
    ReDim PowerSequencePin(PinCnt)

    For Each powerPin In Pins
        TempString = ""
        PowerName = CStr(powerPin)
        
        'get power sequence global spec
        If Power_up_down_flag Then
            TempString = PowerName & "_PowerDownSequence_GLB"
            If TheExec.specs.Globals.Contains(TempString) Then
                PowerSequence = TheExec.specs.Globals(TempString).ContextValue
            Else
                If DebugFlag Then TheExec.Datalog.WriteComment "ERROR:" & PowerName & " does not have power down sequence"
            End If
        Else
            TempString = PowerName & "_PowerSequence_GLB"
            PowerSequence = TheExec.specs.Globals(TempString).ContextValue
        End If
'''        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  debug
'''        If i > 68 Then
'''            PowerSequence = 2
'''        Else
'''            PowerSequence = TheExec.Specs.Globals(TempString).ContextValue
'''        End If
'''        i = i + 1
'''        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If TheExec.DataManager.ChannelType(powerPin) <> "N/C" Then 'check CP or FT NC pins
            ''==============Power down seq which follows Power up seq "PowerSequencePin"==================
            If PowerSequence <> 99 Then 'And Power_up_down_flag Then
                If PowerSequencePin(PowerSequence) = "" Then
                    PowerSequencePin(PowerSequence) = PowerName
                Else
                    PowerSequencePin(PowerSequence) = PowerSequencePin(PowerSequence) & "," & PowerName
                End If
                If PowerSequence >= TempMaxSequence Then TempMaxSequence = PowerSequence
            ''==============Power down seq which follows Power up seq "PowerSequencePin_GB"=================
'            ElseIf Power_up_down_flag = False Then
'                TempMaxSequence = TempMaxSequence_GB
'                PowerSequencePin = PowerSequencePin_GB
            'sequence 99, means disconnect pins
            ElseIf PowerSequence = 99 Then
                pin_name = powerPin
                SlotType = LCase(GetInstrument(pin_name, 0))
                Select Case SlotType
                    Case "dc-07":
                                Vmain = TheHdw.DCVI.Pins(powerPin).Voltage
                                Irange = TheHdw.DCVI.Pins(powerPin).CurrentRange.Value
                                If DebugFlag Then    'debugprint
                                    TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin & "(N/A)", 30, False) & ", Vmain " & Format(Vmain, "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(0, 2, True) & ", RiseTime " & FormatNumericDatalog(0 * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
                                End If
                    Case Else
                                TheHdw.DCVS.Pins(powerPin).Disconnect
                                Vmain = TheHdw.DCVS.Pins(powerPin).Voltage.Main.Value
                                Irange = TheHdw.DCVS.Pins(powerPin).CurrentRange.Value
                                
                                If DebugFlag Then    'debugprint
                                    TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin & "(N/A)", 30, False) & ", Vmain " & Format(Vmain, "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(0, 2, True) & ", FallTime " & FormatNumericDatalog(0 * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
                                End If
                End Select
            Else
            
            End If
        'NC pins, does not need to power off
        Else
            Vmain = 0   'Can not read from DCVS
            Irange = 0
            If DebugFlag Then    'debugprint
                TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin & "(N/C)", 30, False) & ", Vmain " & Format(Vmain, "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(0, 2, True) & ", FallTime " & FormatNumericDatalog(0 * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
            End If
        End If
    Next powerPin
    
    If Power_up_down_flag = False And power_up_en Then  '''follow power up sequence
        TempMaxSequence = TempMaxSequence_GB
        PowerSequencePin = PowerSequencePin_GB
    End If
        
    If power_up_en Then
        For i = TempMaxSequence To 0 Step -1
        
                '------------------------------------Support multiple nWire port 20170503
                For nw_i = 0 To UBound(nwire_seq_GB)
                    If nwire_seq_GB(nw_i) = i And nwire_port_GB(nw_i) <> "" Then
                        TheExec.Datalog.WriteComment vbCrLf & "print: power down for nwire(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                        PowerDown_Interpose nwire_port_GB(nw_i), DebugFlag
                    End If
                Next nw_i
                
                If PowerSequencePin(i) <> "" Then
                    TheExec.Datalog.WriteComment vbCrLf & "print: power down action(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                    PowerOff_I_Meter_Parallel PowerSequencePin(i), WaitConnectTime, WaitConnectTime, i, DebugFlag 'WaitConnectTime, WaitConnectTime, i, DebugFlag
                End If
                '------------------------------------Support I/O init_H
                If io_h_pins <> "" Then
                    For j = 0 To UBound(IO_H_seq_nu_GB)
                        If IO_H_seq_nu_GB(j) = i Then
                            TheHdw.Digital.Pins(IO_H_seq_pin_total_GB(j)).InitState = chInitLo
                            TheExec.Datalog.WriteComment vbCrLf & "print: power down for I/O pins force L : (" & IO_H_seq_pin_total_GB(j) & ") ; sequence :" & j & vbCrLf & RepeatChr("*", 120)
                        End If
                    Next j
                End If
                '------------------------------------Support I/O init_L
                If io_l_pins <> "" Then
                    For j = 0 To UBound(IO_L_seq_nu_GB)
                        If IO_L_seq_nu_GB(j) = i Then
                            TheHdw.Digital.Pins(IO_L_seq_pin_total_GB(j)).InitState = chInitHi
                            TheExec.Datalog.WriteComment vbCrLf & "print: power down for nwire(" & IO_L_seq_pin_total_GB(j) & ")" & vbCrLf & RepeatChr("*", 120)
                            TheExec.Datalog.WriteComment vbCrLf & "print: power down for I/O pins force H : (" & IO_L_seq_pin_total_GB(j) & ") ; sequence :" & j & vbCrLf & RepeatChr("*", 120)
                        End If
                    Next j
                End If
                '------------------------------------Support I/O init_HIZ
                If io_hz_pins <> "" Then
                    For j = 0 To UBound(IO_HZ_seq_nu_GB)
                        If IO_HZ_seq_nu_GB(j) = i Then
                            TheHdw.Digital.Pins(IO_HZ_seq_pin_total_GB(j)).Disconnect
                            TheExec.Datalog.WriteComment vbCrLf & "print: power down for nwire(" & IO_HZ_seq_pin_total_GB(j) & ")" & vbCrLf & RepeatChr("*", 120)
                            TheExec.Datalog.WriteComment vbCrLf & "print: power down for I/O pins force HZ : (" & IO_HZ_seq_pin_total_GB(j) & ") ; sequence :" & j & vbCrLf & RepeatChr("*", 120)
                        End If
                    Next j
                End If
                '------------------------------------
        Next i
    Else
        For i = TempMaxSequence To 0 Step -1
                '------------------------------------Support multiple nWire port 20170503
                For nw_i = 0 To UBound(nwire_pin)
                    If nwire_seq(nw_i) = i Then
                        TheExec.Datalog.WriteComment vbCrLf & "print: power down for nwire(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                        PowerDown_Interpose nwire_pin(nw_i), DebugFlag
                    End If
                Next nw_i
                
                If PowerSequencePin(i) <> "" Then
                    TheExec.Datalog.WriteComment vbCrLf & "print: power down action(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                    PowerOff_I_Meter_Parallel PowerSequencePin(i), WaitConnectTime, WaitConnectTime, i, DebugFlag 'WaitConnectTime, WaitConnectTime, i, DebugFlag
                End If
                '------------------------------------
        Next i
    End If
    
    Call Print_Footer("Power down sequence")
    Exit Function
    
errHandler:
        ErrorDescription ("PowerDown_Parallel")
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PowerUp_Parallel(PowerPinList_DCVS As String, PowerPinList_DCVI As String, IO_H As String, IO_L As String, IO_HZ As String, DisconnectPinList As String, Optional WaitConnectTime As Double = 0.001, Optional DebugFlag As Boolean = False)
'power up sequence at flow start
    Dim CurrentChans As String
    Dim site As Variant
    Dim Pins() As String, PinCnt As Long
    Dim pins_dcvs() As String, PinCnt_dcvs As Long, pins_dcvi() As String, PinCnt_dcvi As Long
    Dim powerPin As Variant
    Dim PowerName As String
    Dim TempString As String
    Dim Vmain As Double
    Dim Irange As Double
    Dim step As Integer
    Dim RiseTime As Double
    Dim PowerSequence As Double
    Dim nwire_port1 As Double
    Dim nwire_port2 As Double
    Dim i As Long, j As Long
    Dim PowerSequencePin() As String
    Dim TempMaxSequence As Long:: TempMaxSequence = 0
    
    Dim nwire_seq() As Long
    Dim nwire_port() As String, nwire_port_plist As New PinList
    Dim nWire_port_ary() As String
    Dim nWire_port_count As Long
    Dim nwp As Variant, nw_i As Long
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim pin_name As String
    Dim SlotType As String
    
    Dim IO_H_seq_nu() As Long, IO_L_seq_nu() As Long, IO_HZ_seq_nu() As Long
    Dim IO_H_seq_pin() As String, IO_L_seq_pin() As String, IO_HZ_seq_pin() As String
    Dim IO_H_seq_pin_total() As String, IO_L_seq_pin_total() As String, IO_HZ_seq_pin_total() As String
    Dim IO_H_list() As String, IO_L_list() As String, IO_HZ_list() As String
    Dim IO_H_count As Long, IO_L_count As Long, IO_HZ_count As Long
    Dim IO_H_nu As Long, IO_L_nu As Long, IO_HZ_nu As Long

    On Error GoTo errHandler
    
    If PowerPinList_DCVS <> "" Then power_dcvs_exit = True
    If PowerPinList_DCVI <> "" Then power_dcvi_exit = True
    
    If IO_H <> "" Then io_h_pins = IO_H
    If IO_L <> "" Then io_l_pins = IO_L
    If IO_HZ <> "" Then io_hz_pins = IO_HZ

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call Print_Header("Power up sequence")
    '===========================================================================================================
    If power_up_en <> True Then
            '''''''''''''''''Support multiple nWire port 20170503'''''''''''''
            nWire_port_ary = Split(nWire_Ports_GLB, ",")
            nWire_port_count = 0
            ReDim nwire_port(UBound(nWire_port_ary) + 1)
            ReDim nwire_seq(UBound(nWire_port_ary) + 1)
            For Each nwp In nWire_port_ary
                Get_nWire_Name CStr(nwp), port_pa, ac_spec_pa, pin_pa, global_spec_pa
                nwire_port(nWire_port_count) = port_pa
                nwire_seq(nWire_port_count) = TheExec.specs.Globals(global_spec_pa).ContextValue
                'If nwire_seq(nWire_port_count) >= TempMaxSequence Then TempMaxSequence = nwire_seq(nWire_port_count)
                nWire_port_count = nWire_port_count + 1
            Next nwp
            '''''''''''''''''Support I/O pins 20180717''''''''''''''''''''''''
            '''''IO_H
            If io_h_pins <> "" Then
                IO_H_nu = 0
                ReDim IO_H_seq_pin_total(100)
                ReDim IO_H_seq_nu(100)
                TheExec.DataManager.DecomposePinList IO_H, IO_H_seq_pin(), IO_H_count
                For i = 0 To IO_H_count - 1
                    IO_H_seq_pin_total(IO_H_nu) = IO_H_seq_pin(i)
                    IO_H_nu = IO_H_nu + 1
                Next i
                ReDim Preserve IO_H_seq_pin_total(IO_H_nu - 1)
                                                                                        
                ReDim Preserve IO_H_seq_nu(UBound(IO_H_seq_pin_total))
                ReDim Preserve IO_H_seq_pin(UBound(IO_H_seq_pin_total))
                                                                                        
                For i = 0 To UBound(IO_H_seq_pin_total)
                    TempString = ""
                    TempString = IO_H_seq_pin_total(i) & "_PowerSequence_GLB_InitHi"
                    IO_H_seq_nu(i) = TheExec.specs.Globals(TempString).ContextValue
                    If IO_H_seq_nu(i) >= TempMaxSequence Then TempMaxSequence = IO_H_seq_nu(i)
                Next i
                ReDim Preserve IO_H_seq_nu(UBound(IO_H_seq_pin_total))
            End If

            '''''IO_L
            If io_l_pins <> "" Then
                IO_L_nu = 0
                ReDim IO_L_seq_pin_total(100)
                ReDim IO_L_seq_nu(100)
                TheExec.DataManager.DecomposePinList IO_L, IO_L_seq_pin(), IO_L_count
                For i = 0 To IO_L_count - 1
                    IO_L_seq_pin_total(IO_L_nu) = IO_L_seq_pin(i)
                    IO_L_nu = IO_L_nu + 1
                Next i
                ReDim Preserve IO_L_seq_pin_total(IO_L_nu - 1)
                                                                                        
                ReDim Preserve IO_L_seq_nu(UBound(IO_L_seq_pin_total))
                ReDim Preserve IO_L_seq_pin(UBound(IO_L_seq_pin_total))
                                                                                        
                For i = 0 To UBound(IO_L_seq_pin_total)
                    TempString = ""
                    TempString = IO_L_seq_pin_total(i) & "_PowerSequence_GLB_initLo"
                    IO_L_seq_nu(i) = TheExec.specs.Globals(TempString).ContextValue
                    If IO_L_seq_nu(i) >= TempMaxSequence Then TempMaxSequence = IO_L_seq_nu(i)
                Next i
                ReDim Preserve IO_L_seq_nu(UBound(IO_L_seq_pin_total))
            End If

            '''''IO_HZ
            If io_hz_pins <> "" Then
                IO_HZ_nu = 0
                ReDim IO_HZ_seq_pin_total(100)
                ReDim IO_HZ_seq_nu(100)
                TheExec.DataManager.DecomposePinList IO_HZ, IO_HZ_seq_pin(), IO_HZ_count
                For i = 0 To IO_HZ_count - 1
                    IO_HZ_seq_pin_total(IO_HZ_nu) = IO_HZ_seq_pin(i)
                    IO_HZ_nu = IO_HZ_nu + 1
                Next i
                ReDim Preserve IO_HZ_seq_pin_total(IO_HZ_nu - 1)
                
                ReDim Preserve IO_HZ_seq_nu(UBound(IO_HZ_seq_pin_total))
                ReDim Preserve IO_HZ_seq_pin(UBound(IO_HZ_seq_pin_total))
                
                For i = 0 To UBound(IO_HZ_seq_pin_total)
                    TempString = ""
                    TempString = IO_HZ_seq_pin_total(i) & "_PowerSequence_GLB_initHZ"
                    IO_HZ_seq_nu(i) = TheExec.specs.Globals(TempString).ContextValue
                    If IO_HZ_seq_nu(i) >= TempMaxSequence Then TempMaxSequence = IO_HZ_seq_nu(i)
                Next i
                ReDim Preserve IO_HZ_seq_nu(UBound(IO_HZ_seq_pin_total))
            End If
    End If
    '===========================================================================================================

    '------------------------------------------------------------------init power pins
    TheExec.Datalog.WriteComment "print: Power up start, Power pins: " & PowerPinList_DCVS & "," & PowerPinList_DCVI
    If PowerPinList_DCVS <> "" Then
        TheHdw.DCVS.Pins(PowerPinList_DCVS).Voltage.Main = 0  'reset to 0V
        TheExec.DataManager.DecomposePinList PowerPinList_DCVS, pins_dcvs(), PinCnt_dcvs
    End If
    If PowerPinList_DCVI <> "" Then
        TheHdw.DCVI.Pins(PowerPinList_DCVI).Voltage = 0
        TheExec.DataManager.DecomposePinList PowerPinList_DCVI, pins_dcvi(), PinCnt_dcvi
    End If
    '------------------------------------------------------------------disconnect I/O pins
    TheHdw.Digital.Pins(DisconnectPinList).Connect
    TheHdw.Digital.Pins(DisconnectPinList).InitState = chInitoff
    
    TheExec.Datalog.WriteComment "print: Power up digital disconnect, Digital pins: " & DisconnectPinList
    TheExec.Datalog.WriteComment RepeatChr("*", 120)
    '------------------------------------------------------------------
    
    PinCnt = PinCnt_dcvs + PinCnt_dcvi
    ReDim Pins(PinCnt - 1)
    
    For i = 0 To PinCnt - 1
        If i < PinCnt_dcvs And PinCnt_dcvs > 0 Then
            Pins(i) = pins_dcvs(i)
        Else
            Pins(i) = pins_dcvi(i - PinCnt_dcvs)
        End If
    Next i
    
    ReDim PowerSequencePin(PinCnt) 'get pin count numbers to arrange array's memory

'''    i = 0    '''for debug
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each powerPin In Pins
        TempString = ""
        PowerName = CStr(powerPin)
        'get power sequence global spec
        TempString = PowerName & "_PowerSequence_GLB"
        PowerSequence = TheExec.specs.Globals(TempString).ContextValue
'''        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''for debug
'''        If i > 68 Then
'''            PowerSequence = 2
'''        Else
'''            PowerSequence = TheExec.Specs.Globals(TempString).ContextValue
'''        End If
'''        i = i + 1
'''        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If TheExec.DataManager.ChannelType(powerPin) <> "N/C" Then
            If PowerSequence <> 99 And power_up_en <> True Then
                'string power sequence pin
                If PowerSequencePin(PowerSequence) = "" Then
                    PowerSequencePin(PowerSequence) = PowerName
                Else
                    PowerSequencePin(PowerSequence) = PowerSequencePin(PowerSequence) & "," & PowerName
                End If
                If PowerSequence >= TempMaxSequence Then TempMaxSequence = PowerSequence
            'sequence 99, means disconnect pins
            ElseIf PowerSequence = 99 Then
                'TheHdw.DCVS.Pins(PowerPin).Disconnect ' it cause voltage spike, removed it
                pin_name = powerPin
                SlotType = LCase(GetInstrument(pin_name, 0))
                Select Case SlotType
                    Case "dc-07":
                                Vmain = TheHdw.DCVI.Pins(powerPin).Voltage
                                Irange = TheHdw.DCVI.Pins(powerPin).CurrentRange.Value
                                If DebugFlag = True Then    'debugprint
                                    TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin & "(N/A)", 30, False) & ", Vmain " & Format(Vmain, "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(0, 2, True) & ", RiseTime " & FormatNumericDatalog(0 * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
                                End If
                    Case Else
                                Vmain = TheHdw.DCVS.Pins(powerPin).Voltage.Main.Value
                                Irange = TheHdw.DCVS.Pins(powerPin).CurrentRange.Value
                                If DebugFlag = True Then    'debugprint
                                    TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin & "(N/A)", 30, False) & ", Vmain " & Format(Vmain, "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(0, 2, True) & ", RiseTime " & FormatNumericDatalog(0 * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
                                End If
                End Select
            Else
            
            End If
        Else
            Vmain = 0                   'Can not read from DCVS
            Irange = 0
            If DebugFlag = True Then    'debugprint
                TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin & "(N/C)", 30, False) & ", Vmain " & Format(Vmain, "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(0, 2, True) & ", RiseTime " & FormatNumericDatalog(0 * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
            End If
        End If
    Next powerPin
    '=========================================================================================Power Sequence
    If power_up_en <> True Then
            ReDim PowerSequencePin_GB(UBound(PowerSequencePin))
            ReDim nwire_seq_GB(nWire_port_count) As Long
            ReDim nwire_port_GB(nWire_port_count) As String
            If io_h_pins <> "" Then
                ReDim IO_H_seq_nu_GB(UBound(IO_H_seq_nu)) As Long
                ReDim IO_H_seq_pin_total_GB(UBound(IO_H_seq_nu)) As String
            End If
            If io_l_pins <> "" Then
                ReDim IO_L_seq_nu_GB(UBound(IO_L_seq_nu)) As Long
                ReDim IO_L_seq_pin_total_GB(UBound(IO_L_seq_nu)) As String
            End If
            If io_hz_pins <> "" Then
                ReDim IO_HZ_seq_nu_GB(UBound(IO_HZ_seq_nu)) As Long
                ReDim IO_HZ_seq_pin_total_GB(UBound(IO_HZ_seq_nu)) As String
            End If

            TempMaxSequence_GB = TempMaxSequence
            
            For i = 0 To TempMaxSequence + 1
                    PowerSequencePin_GB(i) = PowerSequencePin(i)
                    '------------------------------------power pin sequence
                    If PowerSequencePin(i) <> "" Then
                        ''power up
                        TheExec.Datalog.WriteComment vbCrLf & "print: power up action(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                        PowerOn_I_Meter_Parallel PowerSequencePin(i), WaitConnectTime, WaitConnectTime, i, DebugFlag
                    End If
                    '------------------------------------Support multiple nWire port 20170503
                    For nw_i = 0 To nWire_port_count
                        nwire_seq_GB(nw_i) = nwire_seq(nw_i)
                        nwire_port_GB(nw_i) = nwire_port(nw_i)
                        If nwire_seq(nw_i) = i And nwire_port(nw_i) <> "" Then
                            TheExec.Datalog.WriteComment vbCrLf & "print: power up for nwire(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                            nwire_port_plist.Value = nwire_port(nw_i)
                            PowerUp_Interpose nwire_port_plist, DebugFlag
                        End If
                    Next nw_i
                    '------------------------------------Support I/O init_H
                    If io_h_pins <> "" Then
                        For j = 0 To UBound(IO_H_seq_nu)
                            IO_H_seq_nu_GB(j) = IO_H_seq_nu(j)
                            IO_H_seq_pin_total_GB(j) = IO_H_seq_pin_total(j)
                            If IO_H_seq_nu(j) = i And IO_H_seq_pin_total(j) <> "" Then
                                TheHdw.Digital.Pins(IO_H_seq_pin_total(j)).Connect
                                TheHdw.Digital.Pins(IO_H_seq_pin_total(j)).InitState = chInitHi
                                TheExec.Datalog.WriteComment vbCrLf & "print: power up for I/O pins force H : (" & IO_H_seq_pin_total(j) & ") ; sequence :" & i & vbCrLf & RepeatChr("*", 120)
                            End If
                        Next j
                    End If
                    '------------------------------------Support I/O init_L
                    If io_l_pins <> "" Then
                        For j = 0 To UBound(IO_L_seq_nu)
                            IO_L_seq_nu_GB(j) = IO_L_seq_nu(j)
                            IO_L_seq_pin_total_GB(j) = IO_L_seq_pin_total(j)
                            If IO_L_seq_nu(j) = i And IO_L_seq_pin_total(j) <> "" Then
                                TheHdw.Digital.Pins(IO_L_seq_pin_total(j)).Connect
                                TheHdw.Digital.Pins(IO_L_seq_pin_total(j)).InitState = chInitLo
                                TheExec.Datalog.WriteComment vbCrLf & "print: power up for I/O pins force L : (" & IO_L_seq_pin_total(j) & ") ; sequence :" & i & vbCrLf & RepeatChr("*", 120)
                            End If
                        Next j
                    End If
                    '------------------------------------Support I/O init_HIZ
                    If io_hz_pins <> "" Then
                        For j = 0 To UBound(IO_HZ_seq_nu)
                            IO_HZ_seq_nu_GB(j) = IO_HZ_seq_nu(j)
                            IO_HZ_seq_pin_total_GB(j) = IO_HZ_seq_pin_total(j)
                            If IO_HZ_seq_nu(j) = i And IO_HZ_seq_pin_total(j) <> "" Then
                                TheHdw.Digital.Pins(IO_HZ_seq_pin_total(j)).Connect
                                TheHdw.Digital.Pins(IO_HZ_seq_pin_total(j)).InitState = chInitoff
                                TheExec.Datalog.WriteComment vbCrLf & "print: power up for I/O pins force HZ : (" & IO_HZ_seq_pin_total(j) & ") ; sequence :" & i & vbCrLf & RepeatChr("*", 120)
                            End If
                        Next j
                    End If
                    '------------------------------------
            Next i
    Else
            For i = 0 To TempMaxSequence_GB + 1
                    '------------------------------------power pin sequence
                    If PowerSequencePin_GB(i) <> "" Then
                        ''power up
                        TheExec.Datalog.WriteComment vbCrLf & "print: power up action(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                        PowerOn_I_Meter_Parallel PowerSequencePin_GB(i), WaitConnectTime, WaitConnectTime, i, DebugFlag
                    End If
                    '------------------------------------Support multiple nWire port 20170503
                    For nw_i = 0 To UBound(nwire_seq_GB)
                        If nwire_seq_GB(nw_i) = i And nwire_port_GB(nw_i) <> "" Then
                            TheExec.Datalog.WriteComment vbCrLf & "print: power up for nwire(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                            nwire_port_plist.Value = nwire_port_GB(nw_i)
                            PowerUp_Interpose nwire_port_plist, DebugFlag
                        End If
                    Next nw_i
                    '------------------------------------Support I/O init_H
                    If io_h_pins <> "" Then
                        For j = 0 To UBound(IO_H_seq_nu_GB)
                            If IO_H_seq_nu_GB(j) = i And IO_H_seq_pin_total_GB(j) <> "" Then
                                TheHdw.Digital.Pins(IO_H_seq_pin_total_GB(j)).Connect
                                TheHdw.Digital.Pins(IO_H_seq_pin_total_GB(j)).InitState = chInitHi
                                TheExec.Datalog.WriteComment vbCrLf & "print: power up for I/O pins force H : (" & IO_H_seq_pin_total_GB(j) & ") ; sequence :" & i & vbCrLf & RepeatChr("*", 120)
                            End If
                        Next j
                    End If
                    '------------------------------------Support I/O init_L
                    If io_l_pins <> "" Then
                        For j = 0 To UBound(IO_L_seq_nu_GB)
                            If IO_L_seq_nu_GB(j) = i And IO_L_seq_pin_total_GB(j) <> "" Then
                                TheHdw.Digital.Pins(IO_L_seq_pin_total_GB(j)).Connect
                                TheHdw.Digital.Pins(IO_L_seq_pin_total_GB(j)).InitState = chInitLo
                                TheExec.Datalog.WriteComment vbCrLf & "print: power up for I/O pins force L : (" & IO_L_seq_pin_total_GB(j) & ") ; sequence :" & i & vbCrLf & RepeatChr("*", 120)
                            End If
                        Next j
                    End If
                    '------------------------------------Support I/O init_HIZ
                    If io_hz_pins <> "" Then
                        For j = 0 To UBound(IO_HZ_seq_nu_GB)
                            If IO_HZ_seq_nu_GB(j) = i And IO_HZ_seq_pin_total_GB(j) <> "" Then
                                TheHdw.Digital.Pins(IO_HZ_seq_pin_total_GB(j)).Connect
                                TheHdw.Digital.Pins(IO_HZ_seq_pin_total_GB(j)).InitState = chInitoff
                                TheExec.Datalog.WriteComment vbCrLf & "print: power up for I/O pins force HZ : (" & IO_HZ_seq_pin_total_GB(j) & ") ; sequence :" & i & vbCrLf & RepeatChr("*", 120)
                            End If
                        Next j
                    End If
                    '------------------------------------
            Next i
    End If
    '=========================================================================================
    TheHdw.Digital.Pins(DisconnectPinList).Disconnect
    Call Print_Footer("Power up sequence")
    power_up_en = True
    
    Exit Function
    
errHandler:
        ErrorDescription ("PowerUp_Parallel")
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Set_Power_Alarm(Pin_DCVS As PinList, AlarmTime_DCVS As Double, Pin_DCVI As PinList, AlarmTime_DCVI As Double, Pin_GP2 As PinList, AlarmTime_GP2 As Double, Optional b_AlarmForceBin As Boolean = False)
        'add boolean to set force_bin or force_fail

On Error GoTo errHandler
    
''    Pin_DCVS.value = "All_Power"
''    Pin_DCVI.value = "ALL_DCVI"
    
    If AlarmTime_DCVS = 0 Then AlarmTime_DCVS = 0.05
    If AlarmTime_DCVI = 0 Then AlarmTime_DCVI = 0.05
    If AlarmTime_GP2 = 0 Then AlarmTime_GP2 = 0.05
    
    If Pin_DCVS.Value = "" And Pin_DCVI.Value = "" And Pin_GP2.Value = "" Then
        TheExec.Datalog.WriteComment "*************************************************************************"
        TheExec.Datalog.WriteComment "Error on Set_Power_pins_alarm, please fill in pin name to set alarm time."
        TheExec.Datalog.WriteComment "*************************************************************************"
    End If
    '==================================================================
    If Pin_DCVS.Value <> "" Then
        With TheHdw.DCVS.Pins(Pin_DCVS)
            .Gate = False
            TheHdw.Wait 1 * ms '20170209 Add to wait gate off to avoid set ifold time out error
            .mode = tlDCVSModeVoltage
            .Voltage.Main = 0
            
            If b_AlarmForceBin = False Then
                .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail
            Else
                .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin
            End If
            If CStr(AlarmTime_DCVS) = "" Or AlarmTime_DCVS < 0.0001 Then
                TheExec.Datalog.WriteComment "************************************************************"
                TheExec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time."
                TheExec.Datalog.WriteComment "************************************************************"
            End If
            '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime
            .CurrentLimit.Source.FoldLimit.TimeOut = AlarmTime_DCVS
        End With
    End If
    '==================================================================
    If Pin_DCVI.Value <> "" Then
        With TheHdw.DCVI.Pins(Pin_DCVI)
            .Gate = False
            .mode = tlDCVIModeVoltage
            .Voltage = 0
'''            .Alarm(tlDCVIAlarmAll) = tlAlarmForceFail
            If b_AlarmForceBin = False Then
                .Alarm(tlDCVIAlarmAll) = tlAlarmForceFail
            Else
                .Alarm(tlDCVIAlarmAll) = tlAlarmForceBin
            End If
'            .FoldCurrentLimit.TimeOut = AlarmTime             'UVI80 spec 100us ~ 100ms
            .FoldCurrentLimit.TimeOut = AlarmTime_DCVI         '0.05
        End With
    End If
    '==================================================================
    If Pin_GP2.Value = "" Then
    Else
        With TheHdw.DCVS.Pins(Pin_GP2)
            .Gate = False
            TheHdw.Wait 1 * ms '20170209 Add to wait gate off to avoid set ifold time out error
            .mode = tlDCVSModeVoltage
            .Voltage.Main = 0
            
            If b_AlarmForceBin = False Then
                .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail
            Else
                .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin
            End If
            If CStr(AlarmTime_GP2) = "" Or AlarmTime_GP2 < 0.0001 Then
                TheExec.Datalog.WriteComment "*******************************************************************"
                TheExec.Datalog.WriteComment "Error on Set_DCVS_alarm, please put a reasonable alarm time on GP2."
                TheExec.Datalog.WriteComment "*******************************************************************"
            End If
            '.CurrentLimit.Sink.FoldLimit.TimeOut = AlarmTime
            .CurrentLimit.Source.FoldLimit.TimeOut = AlarmTime_GP2
        End With
    End If
    '==================================================================
    Exit Function
 
errHandler:
        ErrorDescription ("Set_Power_pins_alarm")
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Search_UnExistPin() As Long
Dim Pinmap_Sheet, Group_name, All_Groups, Tset_Pins, All_TsetPins, All_PinGroup_Pins As String
Dim cnt, i, j, k, colcnt As Long
Dim Activate_Sheet As Worksheet
Dim Export_sheet As Worksheet
Dim Tset_Sheets() As String
Dim All_Groups_arr() As String
Dim All_PinGroup_Pins_arr() As String
Dim All_TsetPins_arr() As String
Dim PinAry() As String
Dim PinCnt As Long
Dim Not_exist_pins As String
Dim Not_exist_pins_arr() As String

On Error Resume Next

Not_exist_pins = ""
All_Groups = ""
cnt = 0


Set Export_sheet = ThisWorkbook.Sheets("UnExistPin")
If Export_sheet Is Nothing Then
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "UnExistPin"
    Set Export_sheet = ThisWorkbook.Sheets("UnExistPin")
Else
    Application.DisplayAlerts = False
    Sheets("UnExistPin").delete
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "UnExistPin"
    Application.DisplayAlerts = True
    Set Export_sheet = ThisWorkbook.Sheets("UnExistPin")
End If




'Active sheets
For Each Activate_Sheet In ThisWorkbook.Sheets
    If Activate_Sheet.Visible = xlSheetVisible Then
        If LCase(Activate_Sheet.Name) Like "*timeset*" Or LCase(Activate_Sheet.Name) Like "*tsb*" Then
            Activate_Sheet.Activate
            ReDim Preserve Tset_Sheets(cnt)
            Tset_Sheets(cnt) = Activate_Sheet.Name
            cnt = cnt + 1
        ElseIf LCase(Activate_Sheet.Name) Like "*pinmap*" Then
            Activate_Sheet.Activate
            Pinmap_Sheet = Activate_Sheet.Name
        End If
    End If
Next Activate_Sheet


' Parsing all pin groups in pinmap
cnt = Worksheets(Pinmap_Sheet).Cells(4, 3).End(xlDown).row 'row count in pinmap
For i = 4 To cnt
    Group_name = Worksheets(Pinmap_Sheet).Cells(i, 2)
    If Group_name <> "" Then
        If LCase(Group_name) Like "pins_*" Or LCase(Group_name) Like "*_pa*" Then
            If InStr(All_Groups, Group_name) <> 0 Then
            Else
                If All_Groups = "" Then
                    All_Groups = Group_name
                Else
                    All_Groups = All_Groups & "," & Group_name
                End If
            End If
        End If
    End If
Next i

All_Groups_arr = Split(All_Groups, ",")
' Parsing all pins in pin group
For i = 0 To UBound(All_Groups_arr)
    TheExec.DataManager.DecomposePinList All_Groups_arr(i), PinAry(), PinCnt
    For j = 0 To PinCnt - 1
        If InStr(All_PinGroup_Pins, PinAry(j) & ",") = 0 Then
            If All_PinGroup_Pins = "" Then
                All_PinGroup_Pins = PinAry(j) & ","
            Else
                All_PinGroup_Pins = All_PinGroup_Pins & PinAry(j) & ","
            End If
        End If
    Next j
Next i

''''TheExec.AddOutput "*********Search Start*********"
''''TheExec.AddOutput " "
''''TheExec.AddOutput "Search pin Groups:" & All_Groups
''''TheExec.AddOutput "Search Timset Sheets:" & Join(Tset_Sheets, ",")
''''TheExec.AddOutput " "



For i = 1 To UBound(All_Groups_arr) + 2
    If i = 1 Then
        Export_sheet.Cells(i, 1) = "Search pin Groups"
    Else
        Export_sheet.Cells(i, 1) = All_Groups_arr(i - 2)
    End If
Next i
For i = 1 To UBound(Tset_Sheets) + 2
    If i = 1 Then
        Export_sheet.Cells(i, 2) = "Search TimSet Sheets"
    Else
        Export_sheet.Cells(i, 2) = Tset_Sheets(i - 2)
    End If
Next i






colcnt = 3
For i = 0 To UBound(Tset_Sheets)
    All_TsetPins = ""
    cnt = Worksheets(Tset_Sheets(i)).Cells(7, 4).End(xlDown).row 'row count in TimeSet sheet
    For j = 8 To cnt
        Tset_Pins = Worksheets(Tset_Sheets(i)).Cells(j, 4)
        If InStr(All_TsetPins, Tset_Pins & ",") = 0 Then
            If All_TsetPins = "" Then
                All_TsetPins = Tset_Pins & ","
            Else
                All_TsetPins = All_TsetPins & Tset_Pins & ","
            End If
        End If
    Next j
    
    All_TsetPins_arr = Split(All_TsetPins, ",")
    'Search start
    For j = 0 To UBound(All_TsetPins_arr)
        If InStr(All_PinGroup_Pins, All_TsetPins_arr(j) & ",") = 0 Then
            If Not_exist_pins = "" Then
                Not_exist_pins = All_TsetPins_arr(j) & ","
            Else
                Not_exist_pins = Not_exist_pins & All_TsetPins_arr(j) & ","
            End If
        End If
    Next j
    If Not_exist_pins <> "" Then
        Not_exist_pins_arr = Split(Not_exist_pins, ",")
        Not_exist_pins = Mid(Not_exist_pins, 1, Len(Not_exist_pins) - 1)
'''        TheExec.AddOutput " "
'''        TheExec.AddOutput Tset_Sheets(i) & " not exist pins in pinmap:"
'''        TheExec.AddOutput Not_exist_pins
'''        TheExec.AddOutput "----------------------------------------------"
'''        TheExec.AddOutput " "
        
        Export_sheet.Cells(1, colcnt) = "Not exist pins in: " & Tset_Sheets(i)
        For k = 0 To UBound(Not_exist_pins_arr)
            Export_sheet.Cells(k + 2, colcnt) = Not_exist_pins_arr(k)
        Next k
        colcnt = colcnt + 1
        Not_exist_pins = ""
    End If
Next i

TheExec.AddOutput "*********ExportUnExistPins Search End*********"



End Function


Public Function License_Mapping()


Dim version_str As String
Dim tester_path As String
Dim File_path As String
Dim OutputFilePath As String
Dim file_num As Long
Dim tmpStr As String
Dim line_info() As String
ReDim line_info(1)
Dim i As Long
Dim j As Long

Dim Memory_dep_tester_index As Long
Dim Memory_dep_tester_end_index As Long
Dim Speed_tester_index As Long
Dim Speed_tester_end_index As Long


Dim Memory_dep_index As Long
Dim Memory_dep_info() As String

Dim Speed_index As Long

Dim LVM_Fab As Long
Dim Speed_Fab As Long
LVM_Fab = 512
Speed_Fab = 200000000

i = 0
j = 0

If TheExec.Flow.EnableWord("License_check") = True Then
        gL_License_check = 1
End If

If (gL_License_check = 0) Then
    gL_License_check = gL_License_check + 1
    Exit Function
ElseIf gL_License_check = 1 Then
    gL_License_check = 2
Else
    Exit Function
End If



version_str = Trim(Split(CStr(TheExec.SoftwareVersion), "(")(0))
tester_path = "C:\Program Files (x86)\Teradyne\IG-XL\" & version_str & "\tester\"
File_path = "C:\Program Files (x86)\Teradyne\IG-XL\" & version_str & "\tester\LicenseMappingFile_HSD.txt"

OutputFilePath = Dir(File_path)
file_num = FreeFile


Do While OutputFilePath <> ""
    Open File_path For Input As #file_num
    
    Do Until EOF(file_num)
        Line Input #1, tmpStr
        ReDim Preserve line_info(i)
            If InStr(tmpStr, "HSD_MemoryDepth Licenses Assigned:") <> 0 Then
                Memory_dep_tester_index = i
            End If
            
            If InStr(tmpStr, " HSD_MemoryDepth Licenses Assigned:") <> 0 Then
                Memory_dep_tester_index = i + 2
            ElseIf InStr(tmpStr, " HSD_Speed Licenses Assigned:") <> 0 Then
                Speed_tester_index = i + 2
            ElseIf InStr(tmpStr, "VM DEPTHS") <> 0 Then
                Memory_dep_index = i + 2
            ElseIf InStr(tmpStr, "DATARATES") <> 0 Then
                Speed_index = i + 2
            End If
            
            line_info(i) = tmpStr
            i = i + 1
    Loop
    
    For i = Memory_dep_tester_index + 1 To Memory_dep_tester_index + 10
        If line_info(i) = "" Then Memory_dep_tester_end_index = i - 1:: Exit For
    Next i
    
    For i = Speed_tester_index + 1 To Speed_tester_index + 10
        If line_info(i) = "" Then Speed_tester_end_index = i - 1:: Exit For
    Next i
    
    
    
    
    TheExec.Datalog.WriteComment "======== License information start ========" & Chr(10)

    If (Memory_dep_index <> 0 And Speed_index <> 0) Then
        
        TheExec.Datalog.WriteComment "HSD_MemoryDepth Licenses Assigned:"
        For i = Memory_dep_tester_index To Memory_dep_tester_end_index
            TheExec.Datalog.WriteComment "QUANTITY:" & CStr(Trim(Split(line_info(i), Chr(9))(0))) & "   ,ENABLE LEVEL(M license units):" & CStr(Trim(Split(line_info(i), Chr(9))(1)))
        Next i
        TheExec.Datalog.WriteComment ""
        
        
        TheExec.Datalog.WriteComment "HSD_Speed Licenses Assigned:"
        For i = Speed_tester_index To Speed_tester_end_index
            TheExec.Datalog.WriteComment "QUANTITY:" & CStr(Trim(Split(line_info(i), Chr(9))(0))) & "   ,ENABLE LEVEL(Mbps):" & CStr(Trim(Split(line_info(i), Chr(9))(1)))
        Next i
        TheExec.Datalog.WriteComment ""

        TheExec.Flow.TestLimit resultVal:=CDbl(Split(line_info(Memory_dep_index), Chr(9))(1)) + 20, hiVal:=LVM_Fab, Unit:=unitNone, Tname:="Program LVM", ForceResults:=tlForceNone
        TheExec.Flow.TestLimit resultVal:=CDbl(Split(line_info(Speed_index), Chr(9))(0)) * 1000000, hiVal:=Speed_Fab, Unit:=unitHz, Tname:="Program DataRate", scaletype:=scaleMega, ForceResults:=tlForceNone
    
    Else

    'theexec.AddOutput "Please check 'theexec.Licenses.GenerateLicenseRequirementsFile' is set be true and run on-line "
    TheExec.Datalog.WriteComment "Please check 'theexec.Licenses.GenerateLicenseRequirementsFile' is set be true and run on-line "
    End If
    
    TheExec.Datalog.WriteComment "======== License information end ========" & Chr(10)
    
    
    
'Debug.Print OutputFilePath
    Close #file_num
    OutputFilePath = Dir()
Loop

End Function

''20190604AddFunction
Public Function HIP_Init_Datalog_Setup() ''TER190530
    
    TheExec.Datalog.Setup.DatalogSetup.PartResult = True
    TheExec.Datalog.Setup.DatalogSetup.XYCoordinates = True
    TheExec.Datalog.Setup.DatalogSetup.DisableChannelNumberInPTR = True 'disable channel name to stdf, PE's datalog request -- 131225, chihome
    TheExec.Datalog.Setup.DatalogSetup.OutputWidth = 0

    TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = 60
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.testName.Width = 60
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 70

    TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
    
End Function

Public Function Common_UnitTest()

On Error GoTo errHandler
    
    Dim idx As Variant
    Dim long_ As Long
    Dim dec_result_ As Long
    
    Dim dec_ As Double
    Dim dec_ary_() As Double
    Dim dec_result_double_ As Double
    
    Dim bin_ As String
    Dim bin_ary_() As String
    Dim bin_result_ As String
    
    bin_ = "01001111"
    dec_result_ = 79
    long_ = Bin2Dec(bin_)
    If long_ = dec_result_ Then
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform and get the same result " & "Decimal " & long_ & " by Function Bin2Dec"

    Else
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform but do not get the same result " & "Decimal " & long_ & " by Function Bin2Dec"
        
    End If
    
    dec_result_ = 242
    long_ = Bin2Dec_rev(bin_)
    If long_ = dec_result_ Then
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform and get the same result " & "Decimal " & dec_ & " by Function Bin2Dec_rev"

    Else
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform but do not get the same result " & "Decimal " & long_ & " by Function Bin2Dec_rev"
        
    End If
    
    dec_result_ = 242
    dec_ = Bin2Dec_rev_Double(bin_)
    If dec_ = dec_result_ Then
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform and get the same result " & "Double " & dec_ & " by Function Bin2Dec_rev_Double"
      
    Else
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform but do not get the same result " & "Double " & dec_ & " by Function Bin2Dec_rev_Double"
        
    End If
    TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform to " & "Decimal " & dec_ & " by Function Bin2Dec_rev_Double"
    
    dec_result_double_ = 0.9453125
    dec_ = Bin2Dec_rev_Fractional(bin_)
    If dec_ = dec_result_double_ Then
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform and get the same result " & "Fractional " & dec_ & " by Function Bin2Dec_rev_Fractional"

    Else
        TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform but do not get the same result " & "Fractional " & dec_ & " by Function Bin2Dec_rev_Fractional"
        
    End If
    TheExec.Datalog.WriteComment "Binary " & bin_ & " Transform to " & "Decimal " & dec_ & " by Function Bin2Dec_rev_Fractional"
    
    bin_result_ = "011110010"
    bin_ = Dec2BinStr32Bit(8, long_)
    If bin_ = bin_result_ Then
        TheExec.Datalog.WriteComment "Decimal " & long_ & " Transform and get the same result " & "Binary " & bin_ & " by Function Dec2BinStr32Bit"

    Else
        TheExec.Datalog.WriteComment "Decimal " & long_ & " Transform but do not get the same result " & "Binary " & bin_ & " by Function Dec2BinStr32Bit"
        
    End If

    
Exit Function
errHandler:
        TheExec.AddOutput "Error in the VBT Common_UnitTest"
        If AbortTest Then Exit Function Else Resume Next
End Function
