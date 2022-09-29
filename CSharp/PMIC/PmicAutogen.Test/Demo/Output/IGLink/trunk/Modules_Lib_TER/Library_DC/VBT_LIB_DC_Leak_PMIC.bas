Attribute VB_Name = "VBT_LIB_DC_Leak_PMIC"
Option Explicit

Public LeakPinDic As New Dictionary
Public g_LeakIRange As Double

Private Const IO_COND_SPLIT = ","

Public Function DC_IO_Leakage( _
LeakagePatSet As Pattern, _
Measure_Pin_PPMU As String, _
ForceV As String, _
MeasureI_Range As String, _
Optional Validating_ As Boolean _
) As Long

    Dim Site As Variant
    Dim PPMUMeasure As New PinListData
    Dim Pins() As String, Pin_Cnt As Long
    Dim PattArray() As String, PatCount As Long, Pat As String, patt As Variant
    Dim MeasPinArray() As String
    Dim MeasPinGroup() As String
    Dim ForceVArray() As String
    Dim MeasIArray() As String
    Dim TestNum, TestSeq, PinSeq, PinGrpNum As Integer
    Dim LowLimit As Double:: LowLimit = -999
    Dim HiLimit As Double:: HiLimit = 999
    
    If Validating_ Then
        Call PrLoadPattern(LeakagePatSet.Value)
        Exit Function    ' Exit after validation
    End If

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "DC_IO_Leakage"

    '20180628 : check necessary input information
    Call DC_IO_Leakage_InputStringCheck(Measure_Pin_PPMU, ForceV, MeasureI_Range)

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    '20180627 : customized setup before test
    Call DC_IO_Leakage_PreSetup

    If Len(Measure_Pin_PPMU) = 0 Then
        TheExec.Datalog.WriteComment "The Arguments of Measure Pin and ForceV is not match!!!"
        '        Stop
        Exit Function  '// 2019_1213
    End If

    If Len(Measure_Pin_PPMU) = 0 Then
        TheExec.Datalog.WriteComment "The Arguments of Measure Pin and ForceV is not match!!!"
        '        Stop
        Exit Function  '// 2019_1213
    End If

    If Len(Measure_Pin_PPMU) = 0 Then
        TheExec.Datalog.WriteComment "The Arguments of Measure Pin and ForceV is not match!!!"
        '        Stop
        Exit Function  '// 2019_1213
    End If

    If InStr(Measure_Pin_PPMU, IO_COND_SPLIT) > 0 Then
        MeasPinArray = Split(Measure_Pin_PPMU, IO_COND_SPLIT)
        TestNum = UBound(MeasPinArray)
        If InStr(ForceV, IO_COND_SPLIT) > 0 Then
            ForceVArray = Split(ForceV, IO_COND_SPLIT)
        Else
            ReDim ForceVArray(0)
            ForceVArray(0) = ForceV
        End If
        
        If InStr(MeasureI_Range, IO_COND_SPLIT) > 0 Then
            MeasIArray = Split(MeasureI_Range, IO_COND_SPLIT)
        Else
            ReDim MeasIArray(0)
            MeasIArray(0) = MeasureI_Range
        End If
    Else
        TestNum = 0
        ReDim MeasPinArray(0)
        MeasPinArray(0) = Measure_Pin_PPMU
        ReDim ForceVArray(0)
        ForceVArray(0) = ForceV
        ReDim MeasIArray(0)
        MeasIArray(0) = MeasureI_Range
    End If

    If LeakagePatSet.Value <> "" Then
        TheHdw.Patterns(LeakagePatSet).Load
        Call PATT_GetPatListFromPatternSet(LeakagePatSet.Value, PattArray, PatCount)
    Else
        ReDim PattArray(0)
        PattArray(0) = ""
    End If

    For Each patt In PattArray
        If patt <> "" Then
            Pat = CStr(patt)
            Call TheHdw.Patterns(Pat).test(pfAlways, 0)
            Call TheHdw.Digital.Patgen.HaltWait
        End If
    Next patt

    
    For TestSeq = 0 To TestNum
        
        ReDim MeasPinGroup(0)
        MeasPinGroup(0) = MeasPinArray(TestSeq)
        
'20180627 : support to use "+" as delimiter syntax  to separate between test sequences
        For PinGrpNum = 0 To UBound(MeasPinGroup)
        
        'If MeasPinGroup(PinGrpNum) = "LKG_PINS_IO_1p8" Then TheHdw.Utility.Pins("k60").State = tlUtilBitOff
        
        'If MeasPinGroup = "LKG_PINS_IO_1p8" Then TheHdw.Utility.Pins("K60").State = tlUtilBitOff
            Set PPMUMeasure = New PinListData
        
            TheHdw.Digital.Pins(MeasPinGroup(PinGrpNum)).Disconnect
            With TheHdw.PPMU.Pins(MeasPinGroup(PinGrpNum))
                .Gate = tlOff
                .ForceI 0, 0
                .Connect
                .Gate = tlOn
                .ForceV CDbl(ForceVArray(TestSeq)), CDbl(MeasIArray(TestSeq))
            End With
           
            TheExec.DataManager.DecomposePinList MeasPinGroup(PinGrpNum), Pins(), Pin_Cnt
            
            PPMUMeasure.AddPin (MeasPinGroup(PinGrpNum))
            
            TheHdw.Wait 0.005
            
            'perpingrp Setting before measure ( Relay / wait time ... etc )
            
            'Call Leak_PerPinSetting_beforeMeasure(MeasPinGroup(PinGrpNum))


            PPMUMeasure = TheHdw.PPMU.Pins(MeasPinGroup(PinGrpNum)).Read(tlPPMUReadMeasurements, 20)    'normal measure
           
           
            'Call Leak_PerPinSetting_beforeMeasure(MeasPinGroup(PinGrpNum))
            For PinSeq = 0 To PPMUMeasure.Pins.Count - 1
'''                TheExec.Datalog.WriteComment "Pin = " & (PPMUMeasure.Pins(PinSeq) & " Measure Current Range = " & TheHdw.PPMU.Pins(PPMUMeasure.Pins(PinSeq)).MeasureCurrentRange)
                If TheExec.TesterMode = testModeOffline Then
                    For Each Site In TheExec.Sites
                        PPMUMeasure.Pins(PPMUMeasure.Pins(PinSeq)).Value(Site) = Rnd(nA) * nA
                        LowLimit = PPMUMeasure.Pins(PPMUMeasure.Pins(PinSeq)).Value(Site) - 5 * nA
                        HiLimit = PPMUMeasure.Pins(PPMUMeasure.Pins(PinSeq)).Value(Site) + 5 * nA
                    Next Site
                End If
                If UCase(TheExec.DataManager.InstanceName) Like "*HIGH" Then     'Change *HI* -> *HI, Michael
                    DC_CreateTestName "IIH"
                ElseIf UCase(TheExec.DataManager.InstanceName) Like "*LOW" Then
                    DC_CreateTestName "IIL"
                End If
                HiLimit = LL.GetHiLimit(DC_GetTestName)
                LowLimit = LL.GetLowLimit(DC_GetTestName)
                DC_UpdatePinName UCase(PPMUMeasure.Pins(PinSeq))
                LL.FlowTestLimit PPMUMeasure.Pins(PinSeq), LowLimit, HiLimit, scaletype:=scaleNone, Unit:=unitAmp, TName:=DC_GetTestName, forceVal:=Round(TheHdw.PPMU(PPMUMeasure.Pins(PinSeq)).Voltage.Value, 3), ForceUnit:=unitVolt, ForceResults:=tlForceFlow
'                TheExec.Flow.TestLimit PPMUMeasure.Pins(PinSeq), LowLimit, HiLimit, scaletype:=scaleNone, Unit:=unitAmp, TName:=DC_GetTestName, forceVal:=Round(TheHdw.PPMU(PPMUMeasure.Pins(PinSeq)).Voltage.Value, 3), ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            Next PinSeq
            
            With TheHdw.PPMU.Pins(MeasPinGroup(PinGrpNum))
                .Gate = tlOff
                .Disconnect
            End With
             'If MeasPinGroup = "LKG_PINS_IO_1p8" Then TheHdw.Utility.Pins("K60").State = tlUtilBitOn
             
        If MeasPinGroup(PinGrpNum) = "LKG_PINS_IO_1p8" Then TheHdw.Utility.Pins("k60").State = tlUtilBitOn
        
            'Call Leak_PerPinSetting_afterMeasure(MeasPinGroup(PinGrpNum))
          
        Next PinGrpNum
    
                            
    Next TestSeq
    
'print all debug information
    DebugPrintFunc LeakagePatSet.Value
    
'20180627 : customized setup after the test
    Call DC_IO_Leakage_PostSetup
    
Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "error in DC_IO_Leakage"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function

Private Sub DC_IO_Leakage_PreSetup()
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "DC_IO_Leakage_PreSetup"

    'ADG1414_CONTROL &HFF, &H24, &H24, &H24, &H4, &HA4, &H22, &H89, &H22
    
Exit Sub
ErrHandler:
     RunTimeError funcName
     If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub DC_IO_Leakage_PostSetup()
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "DC_IO_Leakage_PostSetup"


Exit Sub
ErrHandler:
     RunTimeError funcName
     If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub DC_IO_Leakage_InputStringCheck( _
Measure_Pin_PPMU As String, _
ForceV As String, _
MeasureI_Range As String)

On Error GoTo ErrHandler
Dim funcName As String:: funcName = "DC_IO_Leakage_InputStringCheck"


    If Len(Measure_Pin_PPMU) = 0 Then
        TheExec.Datalog.WriteComment "[WARNING] DC_IO_Leakage : The Arguments of Measure Pin Should not be Empty!!!"
        '        Stop
        Exit Sub  '// 2019_1213
    ElseIf Len(ForceV) = 0 Then
        TheExec.Datalog.WriteComment "[WARNING] DC_IO_Leakage : The Arguments of Force Voltage Should not be Empty!!!"
        '        Stop
        Exit Sub  '// 2019_1213    ElseIf Len(MeasureI_Range) = 0 Then
        TheExec.Datalog.WriteComment "[WARNING] DC_IO_Leakage : The Arguments of Measure Current Range Should not be Empty!!!"
        '        Stop
        Exit Sub  '// 2019_1213
    End If
        
Exit Sub
ErrHandler:
     RunTimeError funcName
     If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function DC_DCVI_Leakage( _
LeakagePatSet As Pattern, _
Measure_Pin_DCVI As String, _
ForceV As String, _
MeasureI_Range As String, _
Optional Validating_ As Boolean _
) As Long

    Dim Site As Variant
    Dim DCVIMeasure As New PinListData
    Dim Pins() As String, Pin_Cnt As Long
    Dim PattArray() As String, PatCount As Long, Pat As String, patt As Variant
    Dim MeasPinArray() As String
    Dim MeasPinGroup() As String
    Dim ForceVArray() As String
    Dim MeasIArray() As String
    Dim TestNum, TestSeq, PinSeq, PinGrpNum As Integer
    Dim arr_instance() As String:: arr_instance = Split(TheExec.DataManager.InstanceName, "_")
    Dim LowLimit, HiLimit As Double
    
    If Validating_ Then
        Call PrLoadPattern(LeakagePatSet.Value)
        Exit Function    ' Exit after validation
    End If

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "DC_DCVI_Leakage"

    '20180628 : check necessary input information
    Call DC_IO_Leakage_InputStringCheck(Measure_Pin_DCVI, ForceV, MeasureI_Range)

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    '20180627 : customized setup before test
    Call DC_IO_Leakage_PreSetup

    If Len(Measure_Pin_DCVI) = 0 Then
        TheExec.Datalog.WriteComment "The Arguments of Measure Pin and ForceV is not match!!!"
        ''        Stop
        Exit Function  '// 2019_1213
    End If

    If InStr(Measure_Pin_DCVI, IO_COND_SPLIT) > 0 Then
        MeasPinArray = Split(Measure_Pin_DCVI, IO_COND_SPLIT)
        TestNum = UBound(MeasPinArray)
        If InStr(ForceV, IO_COND_SPLIT) > 0 Then
            ForceVArray = Split(ForceV, IO_COND_SPLIT)
        Else
            ReDim ForceVArray(TestNum)
            Dim i As Integer
            For i = 0 To TestNum
                ForceVArray(i) = ForceV
            Next i
        End If
        
        If InStr(MeasureI_Range, IO_COND_SPLIT) > 0 Then
            MeasIArray = Split(MeasureI_Range, IO_COND_SPLIT)
        Else
            ReDim MeasIArray(0)
            MeasIArray(0) = MeasureI_Range
        End If
    Else
        TestNum = 0
        ReDim MeasPinArray(0)
        MeasPinArray(0) = Measure_Pin_DCVI
        ReDim ForceVArray(0)
        ForceVArray(0) = ForceV
        ReDim MeasIArray(0)
        MeasIArray(0) = MeasureI_Range
    End If

    If LeakagePatSet.Value <> "" Then
        TheHdw.Patterns(LeakagePatSet).Load
        Call PATT_GetPatListFromPatternSet(LeakagePatSet.Value, PattArray, PatCount)
    Else
        ReDim PattArray(0)
        PattArray(0) = ""
    End If

    For Each patt In PattArray
        If patt <> "" Then
            Pat = CStr(patt)
            Call TheHdw.Patterns(Pat).test(pfAlways, 0)
            Call TheHdw.Digital.Patgen.HaltWait
        End If
    Next patt
                   
    
    For TestSeq = 0 To TestNum
            ReDim MeasPinGroup(0)
            MeasPinGroup(0) = MeasPinArray(TestSeq)


'20180627 : support to use "+" as delimiter syntax  to separate between test sequences
        For PinGrpNum = 0 To UBound(MeasPinGroup)

            Set DCVIMeasure = New PinListData
            
            TheHdw.DCVI.Pins(MeasPinGroup(PinGrpNum)).Disconnect
            
            If MeasPinGroup(PinGrpNum) Like "*UVI80*" Then
                With TheHdw.DCVI.Pins(MeasPinGroup(PinGrpNum))
                    .Gate = tlOff
                    .Mode = tlDCVIModeVoltage
                    .SetVoltageAndRange CDbl(ForceVArray(TestSeq)), 7
                    .SetCurrentAndRange 20 * uA, 20 * uA
                    .Meter.Mode = tlDCVIMeterCurrent
                    .Meter.CurrentRange = CDbl(MeasIArray(TestSeq))
                    .Connect
                    .BleederResistor = tlDCVIBleederResistorOff
                    .VoltageRange.AutoRange = False
                    .Gate = tlOn
                    
                End With
            ElseIf MeasPinGroup(PinGrpNum) Like "*DC30*" Then
'                With TheHdw.DCVI.Pins(MeasPinGroup(PinGrpNum))
'                    .Gate = tlOff
'                    .Mode = tlDCVIModeVoltage
'                    .ComplianceRange(tlDCVICompliancePositive) = 30
'                    .ComplianceRange(tlDCVIComplianceNegative) = 30
'                    .SetVoltageAndRange CDbl(ForceVArray(TestSeq)), 30 ' IDAC_OUT need to be forced 24V
'                    .SetCurrentAndRange 20 * uA, 20 * uA
'                    .Meter.Mode = tlDCVIMeterCurrent
'                    .Meter.CurrentRange = CDbl(MeasIArray(TestSeq))
'                    .Connect
'                    .VoltageRange.Autorange = False
'                    .Gate = tlOn
'                End With
                
                With TheHdw.DCVI.Pins(MeasPinArray(PinGrpNum))
                    .Gate = tlOff
                    .Mode = tlDCVIModeVoltage
                    .ComplianceRange(tlDCVICompliancePositive) = 30
                    .ComplianceRange(tlDCVIComplianceNegative) = 30
                    .SetVoltageAndRange CDbl(ForceVArray(PinGrpNum)), 30 ' IDAC_OUT need to be forced 24V
                    .SetCurrentAndRange 20 * uA, 20 * uA
                    .Meter.Mode = tlDCVIMeterCurrent
                    .Meter.CurrentRange = CDbl(MeasIArray(PinGrpNum))
                    .Connect
                    .VoltageRange.AutoRange = False
                    .Gate = tlOn
                    
                End With
            Else
                Debug.Print "Measure pin connect neither UVI80 nor DC30"
            End If

            
            DCVIMeasure.AddPin (MeasPinGroup(PinGrpNum))
            
            'TheHdw.Wait 0.02
            
            
            'perpingrp Setting before measure ( Relay / wait time ... etc )
            
            'Call Leak_PerPinSetting_beforeMeasure(MeasPinGroup(PinGrpNum))
            
            DCVIMeasure = TheHdw.DCVI.Pins(MeasPinGroup(PinGrpNum)).Meter.Read(tlStrobe, 10, 100 * KHz, tlDCVIMeterReadingFormatAverage)
            
            For PinSeq = 0 To DCVIMeasure.Pins.Count - 1
'''                TheExec.Datalog.WriteComment "Pin = " & (DCVIMeasure.Pins(PinSeq) & " Measure Current Range = " & TheHdw.PPMU.Pins(DCVIMeasure.Pins(PinSeq)).MeasureCurrentRange)
                If TheExec.TesterMode = testModeOffline Then
                    For Each Site In TheExec.Sites
                        DCVIMeasure.Pins(DCVIMeasure.Pins(PinSeq)).Value(Site) = Rnd(nA) * nA
                        LowLimit = DCVIMeasure.Pins(DCVIMeasure.Pins(PinSeq)).Value(Site) - 5 * nA
                        HiLimit = DCVIMeasure.Pins(DCVIMeasure.Pins(PinSeq)).Value(Site) + 5 * nA
                    Next Site
                End If
                
                If UCase(arr_instance(UBound(arr_instance))) Like "*HI*" Then
                    DC_CreateTestName "IIH"
                ElseIf UCase(arr_instance(UBound(arr_instance))) Like "*LO*" Then
                    DC_CreateTestName "IIL"
                End If
                
                DC_UpdatePinName UCase(DCVIMeasure.Pins(PinSeq)), Format(ForceVArray(TestSeq), "0.000") + "V"
                
  
                
                HiLimit = LL.GetHiLimit(DC_GetTestName)
                LowLimit = LL.GetLowLimit(DC_GetTestName)
                LL.FlowTestLimit ResultVal:=DCVIMeasure.Pins(PinSeq), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=Round(TheHdw.DCVI.Pins(DCVIMeasure.Pins(PinSeq)).Voltage, 3), ForceUnit:=unitVolt, ForceResults:=tlForceFlow
'                TheExec.Flow.TestLimit ResultVal:=DCVIMeasure.Pins(PinSeq), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=Round(TheHdw.DCVI.Pins(DCVIMeasure.Pins(PinSeq)).Voltage, 3), ForceUnit:=unitVolt, ForceResults:=tlForceFlow

            Next PinSeq
            
            With TheHdw.DCVI.Pins(MeasPinGroup(PinGrpNum))
                .Gate = tlOff
                .Disconnect
            End With
            
            'Call Leak_PerPinSetting_afterMeasure(MeasPinGroup(PinGrpNum))
            
        Next PinGrpNum
        
        
        
    Next TestSeq
    
'print all debug information
    DebugPrintFunc LeakagePatSet.Value
    
'20180627 : customized setup after the test
    '''Call DC_IO_Leakage_PostSetup
    
Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "error in DC_IO_Leakage"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function




