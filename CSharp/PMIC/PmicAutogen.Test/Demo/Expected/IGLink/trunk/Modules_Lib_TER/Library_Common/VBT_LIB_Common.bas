Attribute VB_Name = "VBT_LIB_Common"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit
'Revision History:
'V0.0 initial bring up
'V0.1 add keep alive function
'V0.2 add disable compare and enable compare function.
'variable declaration
Public Const Version_Lib_VBT_Common = "0.1"  'lib version

Public DicDiffPairs As New Scripting.Dictionary  'relocation for minimum VBT with RF code'*****************************************



'*****************************************
'******         free run clk, nWire ******
'*****************************************


Public Function StartSBClock(SBFreq As Double) As Long
    On Error GoTo ErrHandler
    Dim SBC_Enable As Long
    'Dim SBFreq As Double
    'TheExec.Datalog.WriteComment "******************  Enable Support BD clock ****************"
    'SBFreq = TheExec.specs.Globals("SBC_Freq_Var").ContextValue
    
    With TheHdw.DIB.SupportBoardClock
        .Connect
        .Frequency = SBFreq
        .VIH = XI0_ref_VOH ' Max is 6V
        .VIL = 0 ' Min is -1V
        .Start
    End With
    SBC_Enable = 1
    TheExec.Flow.TestLimit SBC_Enable, 1, 1, tlSignGreaterEqual, tlSignLessEqual, TName:="SBC enable" 'BurstResult=1:Pass
    'printing in data log
    TheExec.Datalog.WriteComment "********** support board clock = " & Format(SBFreq / 1000000, "0.000") & " Mhz, Clock_Vih = " _
                             & XI0_ref_VOH & " V, Clock_Vil = " & XI0_ref_VOL & " V  *******"
    
    Exit Function
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function StopSBClock()
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "DSPinfo"

    Dim SBC_Enable As Long
    ' Stop and disconnect the support board clock.
    With TheHdw.DIB.SupportBoardClock
        .Stop
        .Disconnect
    End With
    SBC_Enable = 0
    TheExec.Flow.TestLimit SBC_Enable, 1, 1, tlSignGreaterEqual, tlSignLessEqual, TName:="SBC disable" 'BurstResult=1:Pass
    'printing in data log
    TheExec.Datalog.WriteComment "******************  Disable Support BD clock ****************"
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function




Function Start_Profile(PinName As PinList, WhatToCapture As String, SampleRate As Double, SampleSize As Long, Optional CapSignalName As String = "Capture_signal")
'start current or voltage profile capturing

On Error GoTo ErrHandler
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
            .Mode = tlDCVSMeterCurrent
            .Range = TheHdw.DCVS.Pins(PinName).CurrentRange.max '2
        Else
            .Mode = tlDCVSMeterVoltage
            .Range = 10
        End If
        .SampleRate = SampleRate
        .SampleSize = SampleSize
    
    End With
    
    ' Setup the hardware by loading the signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName).LoadSettings
    
    ' Start the capture
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName).Trigger

    Exit Function
ErrHandler:
    ErrorDescription ("Start_Profile")
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function start_profile_DCVI(PinName As String, WhatToCapture As String, SampleRate As Double, SampleSize As Long)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "start_profile_DCVI"

Do While TheHdw.DCVI.Pins(PinName).Capture.IsCaptureDone = False        ' Wait if another capture is running
Loop
TheHdw.DCVI.Pins(PinName).Capture.Signals.Add "Capture_signal"              'Create a SIGNAL to set up instrument
TheHdw.DCVI.Pins(PinName).Capture.Signals.DefaultSignal = "Capture_signal"  'Set this as the default signal

        With TheHdw.DCVI.Pins(PinName)
            .Gate = False
            .Mode = tlDCVIModeCurrent
            .Voltage = 6
            .VoltageRange.AutoRange = True
            .CurrentRange.AutoRange = True
            .Current = 0
            .Connect tlDCVIConnectDefault
            .Gate = True
        End With
 
With TheHdw.DCVI.Pins(PinName).Capture.Signals.Item("Capture_signal")    ' Define the signal used for the capture
    .Reinitialize
    If (WhatToCapture = "I") Then
        .Mode = tlDCVIMeterCurrent
        .Range = 0.02
    Else
        .Mode = tlDCVIMeterVoltage
         .Range = 7
    End If
    .SampleRate = SampleRate
    .SampleSize = SampleSize
End With

TheHdw.DCVI.Pins(PinName).Capture.Signals.Item("Capture_signal").LoadSettings  ' Setup the hardware by loading the signal
TheHdw.DCVI.Pins(PinName).Capture.Signals.Item("Capture_signal").Trigger            ' Start the capture
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Plot_Profile(PinName As PinList, Optional CapSignalName As String = "Capture_signal", Optional ExportWaveform As Boolean = False)
'Plot profiles

    Dim DSPW As New DSPWave
    Dim Label As String
    Dim Site As Variant
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
    
    On Error GoTo ErrHandler

    Do While TheHdw.DCVS.Pins(PinName).Capture.IsRunning = True
    Loop

    day_code = CStr(Year(Now)) & Right("0" & CStr(Month(Now)), 2) & Right("0" & CStr(Day(Now)), 2)
    day_code = day_code & Right("0" & CStr(Hour(Now)), 2) & Right("0" & CStr(Minute(Now)), 2) & Right("0" & CStr(Second(Now)), 2)
    ' Get the captured samples from the instrument
    Call TheExec.DataManager.DecomposePinList(PinName, Pin_Ary(), Pin_Cnt)
    Dim sampleR As String
    For Each p In Pin_Ary
        If TheExec.DataManager.ChannelType(p) <> "N/C" Then
            DSPW = TheHdw.DCVS.Pins(p).Capture.Signals(CapSignalName).DSPWave
            For Each Site In TheExec.Sites
'                TheHdw.Digital.Patgen.ReadLastStart lastBurstPat, isGrp, lastLabel
                 sampleR = CStr(TheHdw.DCVS.Pins(p).Capture.SampleRate)
                'If thehdw.DCVS.Pins(p).Meter.mode = thehdw.DCVS.Pins(p).CurrentRange.Max Then
                If TheHdw.DCVS.Pins(p).Meter.Mode = tlDCVSMeterCurrent Then
                    Label = "Current Profile for Site: " & Site & " " & " " & CapSignalName & "Pin :" & " " & p
                    FileName = "CurrentProfile-Site" & Site & "-" & p & "-" & sampleR & "-" & Current_Insatance & "_" & day_code & ".txt"
                Else
                    Label = "Voltage Profile for Site: " & Site & " " & " " & CapSignalName & "Pin :" & " " & p
                    FileName = "VoltageProfile-Site" & Site & "-" & p & "-" & sampleR & "-" & Current_Insatance & "_" & day_code & ".txt"
                End If
                
                If True Then DSPW.Plot Label   'for pliot
                If ExportWaveform Then
                    Dim TempStr As String
                    TempStr = "D:\" & p
                    Dim fso As New FileSystemObject
                     
                    If Dir(TempStr, vbDirectory) = Empty Then
                        MkDir TempStr
                    End If
                    DSPW.FileExport TempStr & "\" & FileName, File_txt
                End If
                If LCase(GetInstrument(CStr(p), 0)) <> "hexvs" Then DSPW.Clear
            Next Site
        End If
    Next p
    m1_InstanceName = ""
    Exit Function
ErrHandler:
    ErrorDescription ("Plot_Profile")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Plot_profile_DCVI(PinName As String)

Dim DSPW As New DSPWave
Dim Label As String
Dim Site As Variant

On Error GoTo ErrHandler

' Get the captured samples from the instrument
DSPW = TheHdw.DCVI.Pins(PinName).Capture.Signals.Item("Capture_signal").DSPWave

For Each Site In TheExec.Sites.Active
    If TheHdw.DCVI.Pins(PinName).Meter.Mode = tlDCVIMeterCurrent Then
        Label = "Current Profile for Site: " & Site
    Else
        Label = "Voltage Profile for Site: " & Site
End If
    
     DSPW.Plot Label
    
Next Site

Exit Function

ErrHandler:
        TheExec.AddOutput "Error in the Plot Profile"
                If AbortTest Then Exit Function Else Resume Next
End Function

Function Start_Profile_AutoResolution(PinName As String, WhatToCapture As String, SampleRate As Double, SampleSize As Long, Optional CapSignalName As String = "Capture_signal", Optional Plottime As Double = 0)
'start current or voltage profile capturing

On Error GoTo ErrHandler
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
                .Mode = tlDCVSMeterCurrent
                .Range = TheHdw.DCVS.Pins(PinName).CurrentRange.max '2
            Else
                .Mode = tlDCVSMeterVoltage
                .Range = 10
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

ErrHandler:
    ErrorDescription ("Start_Profile_AutoResolution")
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Print_Footer(PrintInfo As String)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "DSPinfo"

    TheExec.Datalog.WriteComment "******************************"
    TheExec.Datalog.WriteComment "*print: " & PrintInfo & " end*"
    TheExec.Datalog.WriteComment "******************************"

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Print_Header(PrintInfo As String)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Print_Header"

    TheExec.Datalog.WriteComment "********************************"
    TheExec.Datalog.WriteComment "*print: " & PrintInfo & " start*"
    TheExec.Datalog.WriteComment "********************************"

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
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

Public Function Print_PgmInfo()
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "printHeaederInfo"

    Call HEADERINFO ' Get the flow control words, calc a hash and record it to datalog/stdf
      
 Exit Function
ErrHandler:
    'Call TheExec.AddOutput("VBT_HDRobj encountered an error with STDTestOnProgStart.  More Info:" & vbCrLf & err.Description, vbBlue, False)
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'*****************************************
'******            Read/Write EPPROM******
'*****************************************
Public Function Write_DIB_EEPROM(Optional DIB_SerialNumber As String) As Long
    On Error GoTo ErrHandler
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
ErrHandler:
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
    On Error GoTo ErrHandler
    
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
ErrHandler:
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
On Error GoTo ErrHandler
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
        TheExec.Flow.TestLimit Prober_Temp, lowVal:=Temp_Lolimit, hiVal:=Temp_Hilimit, TName:="Prober_Temp_" & CStr(Mid(TheExec.DataManager.InstanceName, 16, 18))
        'TheExec.Datalog.WriteComment "/*********************************************/"
       
    Exit Function
ErrHandler:
        TheExec.Datalog.WriteComment "Read Prober Temp VBT function is error "
        TheExec.Datalog.WriteComment "Registry String :" & RegKeyRead("Prober_Temp")
        TheExec.Datalog.WriteComment ("Error #: " & Str(err.Number) & " " & err.Description)
        If AbortTest Then Exit Function Else Resume Next
End Function



Public Function Read_Package_ID()

Dim Site As Variant

On Error GoTo ErrHandler

        TheExec.Datalog.WriteComment "********************************"
        
        For Each Site In TheExec.Sites
            TheExec.Datalog.WriteComment ("Site:" & Site & " Device ID : " & RegKeyRead("ManualTestDeviceID"))
        Next Site
        
        TheExec.Datalog.WriteComment "********************************"
        
        Dim TestName As String
        TestName = "Device ID"
        Dim ResultVal_DeviceID As Double
        ResultVal_DeviceID = CDbl(RegKeyRead("ManualTestDeviceID"))
        For Each Site In TheExec.Sites
            
            TheExec.Flow.TestLimit ResultVal:=ResultVal_DeviceID, TName:=TestName, ForceResults:=tlForceNone
        Next Site
        
       
    Exit Function
ErrHandler:
        TheExec.Datalog.WriteComment "Read Package ID error "
        If AbortTest Then Exit Function Else Resume Next
End Function



Public Function printHeaederInfo()
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "printHeaederInfo"
    Call HEADERINFO
    
    
 Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function Read_TesterConfig()  'ReadFromFile
 Dim funcName As String:: funcName = "Read_TesterConfig"
  On Error GoTo ErrHandler
    Dim i As Integer
    Dim FilePath As String
    Dim LineFromFile As String
    Dim lSlotNum As String
    Dim slotinfo(25) As String
    Dim recordflag As Boolean: recordflag = False
    FilePath = "C:\Program Files (x86)\Teradyne\IG-XL\9.00.00_uflx\tester\CurrentConfig.txt"
    
    Open FilePath For Input As #1
    For i = 0 To 25
    slotinfo(i) = "            "
    Next i
    Do Until EOF(1)
    

    Line Input #1, LineFromFile
    If LineFromFile Like "*slot*" Then
    recordflag = True
    End If
    If recordflag = True Then
        lSlotNum = Left(LineFromFile, 4)
         If IsNumeric(lSlotNum) Then
                
                If LineFromFile Like "*DC-07*" Then
                'judge UVI80/UVI80+
                    If LineFromFile Like "*604-375-02*" Then
                       If CLng(lSlotNum) And 1 Then
                           slotinfo(CLng(lSlotNum)) = "    UVI80   "
                       Else
                           slotinfo(CLng(lSlotNum)) = "   UVI80    "
                       End If
                    End If
                    If LineFromFile Like "*604-375-12*" Then
                       If CLng(lSlotNum) And 1 Then
                           slotinfo(CLng(lSlotNum)) = "   UVI80+   "
                       Else
                           slotinfo(CLng(lSlotNum)) = "   UVI80+   "
                       End If
                    End If
                    
                
                ElseIf LineFromFile Like "*SupportBoard*" Then
                    slotinfo(CLng(lSlotNum)) = "SupportBoard"
                    
    
                ElseIf LineFromFile Like "*DC-30*" Then
                
                      If CLng(lSlotNum) And 1 Then
                           slotinfo(CLng(lSlotNum)) = "    DC-30   "
                       Else
                           slotinfo(CLng(lSlotNum)) = "   DC-30    "
                       End If
                    
                ElseIf LineFromFile Like "*HSD-U*" Then
                       If CLng(lSlotNum) And 1 Then
                           slotinfo(CLng(lSlotNum)) = "   UP1600   "
                       Else
                           slotinfo(CLng(lSlotNum)) = "   UP1600   "
                       End If
                       
                ElseIf LineFromFile Like "*UltraPAC*" Then
                
                      If CLng(lSlotNum) And 1 Then
                           slotinfo(CLng(lSlotNum)) = " UltraPAC   "
                       Else
                           slotinfo(CLng(lSlotNum)) = "  UltraPAC "
                       End If
                End If
       
        End If
   
    End If
    Loop
    
    Close #1

    TheExec.Datalog.WriteComment "==============================Slot Information================================"
    TheExec.Datalog.WriteComment "|25|" & slotinfo(25) & "----------------------|----------------------" & slotinfo(22) & "|22|"
    TheExec.Datalog.WriteComment "|21|" & slotinfo(21) & "----------------------|----------------------" & slotinfo(18) & "|18|"
    TheExec.Datalog.WriteComment "|17|" & slotinfo(17) & "----------------------|----------------------" & slotinfo(14) & "|14|"
    TheExec.Datalog.WriteComment "|13|" & slotinfo(13) & "----------------------|----------------------" & slotinfo(10) & "|10|"
    TheExec.Datalog.WriteComment "|09|" & slotinfo(9) & "----------------------|----------------------" & slotinfo(6) & "|06|"
    TheExec.Datalog.WriteComment "|05|" & slotinfo(5) & "----------------------|----------------------" & slotinfo(2) & "|02|"
    TheExec.Datalog.WriteComment "|01|" & slotinfo(1) & "----------------------|----------------------" & slotinfo(0) & "|00|"
    TheExec.Datalog.WriteComment "|03|" & slotinfo(3) & "----------------------|----------------------" & slotinfo(4) & "|04|"
    TheExec.Datalog.WriteComment "|07|" & slotinfo(7) & "----------------------|----------------------" & slotinfo(8) & "|08|"
    TheExec.Datalog.WriteComment "|11|" & slotinfo(11) & "----------------------|----------------------" & slotinfo(12) & "|12|"
    TheExec.Datalog.WriteComment "|15|" & slotinfo(15) & "----------------------|----------------------" & slotinfo(16) & "|16|"
    TheExec.Datalog.WriteComment "|19|" & slotinfo(19) & "----------------------|----------------------" & slotinfo(20) & "|20|"
    TheExec.Datalog.WriteComment "|23|" & slotinfo(23) & "----------------------|----------------------" & slotinfo(24) & "|24|"
    TheExec.Datalog.WriteComment "=============================================================================="
     
    Exit Function
    
ErrHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function




Public Sub RunTimeError(funcName As String)
    ' Sanity clause
    If TheExec Is Nothing Then

        '//2019_1213
        '        MsgBox "IG-XL in not running!  Error encountered in Exec Interpose Function " + funcName + vbCrLf + _
                 '            "VBT Error # " + Trim$(CStr(err.Number)) + ": " + err.Description
        TheExec.Datalog.WriteComment "IG-XL in not running!  Error encountered in Exec Interpose Function " + funcName + vbCrLf + _
                                     "VBT Error # " + Trim$(CStr(err.Number)) + ": " + err.Description
        Exit Sub
    End If
    TheExec.Datalog.WriteComment "Error encountered in Function::" + funcName
End Sub

