Attribute VB_Name = "LIB_Common_DSSC_HRAM"
Option Explicit
'Revision History:
'V0.0 initial bring up
'*****************************************
'******                         DSSC******
'*****************************************

''''20170920 update for the mutiple Src cases as CFG_RAW+MSP
Public Function DSSC_SetupDigSrcWave(patt As String, DigSrcPin As PinList, SignalName As String, SegmentSize As Long, WaveDefArray() As Long)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "DSSC_SetupDigSrcWave"
    
    'store efuse program bit into a DSP wave
    Dim InWave As New DSPWave
    Dim site As Variant
    Dim WaveDef As String
    
    ''WaveDef = "WaveDef" ''''was
    InWave.Data = WaveDefArray
    site = TheExec.sites.SiteNumber
    ''''20170920 <NOTICE> if multiple apply this function call/sequence to avoid the following SrcWave to overwrite the previous one
    WaveDef = "WaveDef_" + SignalName + "_" & site
    'TheHdw.Patterns(patt).Load
    TheExec.WaveDefinitions.CreateWaveDefinition WaveDef, InWave, True
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName
    With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
        .WaveDefinitionName = WaveDef
        .SampleSize = SegmentSize
        .Amplitude = 1
        .LoadSamples
        .LoadSettings
    End With
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

'*****************************************
'******                         HRAM******
'*****************************************
Function DatalogHRAMVecNum(excludePreTrig As Boolean, StvNum As Long, DataOutputPins As PinList)

    Dim maxIdx As Long
    Dim idx As Long
    Dim preTrig As Long
    preTrig = 0
    Dim PinData As New PinListData
    Dim Binary_String As String

    '=== Simulated Data ===
    If (TheExec.TesterMode = testModeOffline) Then
       ' PinData = thehdw.Digital.pins(DataOutputPins).HRAM.PinData(0, 1, StvNum)
        With TheHdw.Digital.HRAM
            If excludePreTrig = True Then
                preTrig = .PreTrigCycles
            End If
            'TheExec.Datalog.WriteComment "  Pattern: " + CStr(.PatGenInfo(idx, pgPattern))
    
            Binary_String = ""
            maxIdx = StvNum - 1
            For idx = preTrig To maxIdx
                TheExec.Datalog.WriteComment "      Hram index:" + CStr(idx) + _
                " Vector number: " + CStr(idx) + "   DUT state : " + "0" '+ PinData.pins(DataOutputPins).Value(0)(idx)
                Binary_String = Binary_String + "0"
            Next
        End With
    Else

  
        PinData = TheHdw.Digital.Pins(DataOutputPins).HRAM.PinData(0, 1, StvNum)
        With TheHdw.Digital.HRAM
            If excludePreTrig = True Then
                preTrig = .PreTrigCycles
            End If
            TheExec.Datalog.WriteComment "  Pattern: " + CStr(.PatGenInfo(idx, pgPattern))
    
            Binary_String = ""
            maxIdx = .CapturedCycles - 1
            For idx = preTrig To maxIdx
                TheExec.Datalog.WriteComment "      Hram index:" + CStr(idx) + _
                " Vector number: " + CStr(.PatGenInfo(idx, pgVector)) + "   DUT state : " + PinData.Pins(DataOutputPins).Value(0)(idx)
                If LCase(PinData.Pins("tdo").Value(0)(idx)) = "l" Then
                    Binary_String = Binary_String + "0"
                Else
                    Binary_String = Binary_String + "1"
                End If
            Next
        End With
    

    End If
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbTab + "Binary Code = " + Binary_String

End Function
Public Function Hram_Trig_Setup(TrigType As TrigType, CaptType As CaptType)
    
    Dim WaitForEvent As Boolean
    Dim preTrigCnt As Integer
    Dim stopFull As Boolean

    WaitForEvent = False
    preTrigCnt = 0
    stopFull = True
    With TheHdw.Digital.HRAM
        .SetTrigger TrigType, WaitForEvent, preTrigCnt, stopFull
        .CaptureType = CaptType  'the vector to be captured only at stv micro-code
        .Size = 0
    End With

End Function

