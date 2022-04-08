Attribute VB_Name = "VBT_LIB_Digital_Mbist"
Option Explicit
'Revision History:
'V0.0 initial bring up

Public Type MBIST_BLOCK_TYPE
    MemArray() As Long
    MemStrArray() As String
    MaxRow As Long
    strServerName() As String
    strDecsName() As String
End Type

Public MbistBlock(10) As MBIST_BLOCK_TYPE
Public MbistBlockName

Public tpEvaPattCycleBlockInfor() As EvaPattMbistCycleBlock

Private Type EvaPattMbistCycleBlock
    strBlaclName As String
    lVector As Long
    lCycle As Long
    strCompare As String
    strFlagName As String   '' Add for MbistFP Binning          201606014 Webster
End Type


Private Type MbistCycleBlock
    strPattName As String
    tpMbistCycleBlock() As EvaPattMbistCycleBlock
    strServerName() As String
    strDecsName() As String
    Dic_VectorIdx As New Dictionary
    Dic_CycleIdx As New Dictionary
End Type

Public tpCycleBlockInfor() As MbistCycleBlock

Public gl_MbistFP_Binout As Boolean      ''201606014 webster
Dim gl_strFlagArr() As String
Private Type FlagInfo
    flagName As String
    CheckInfo As Boolean
End Type
Public tyFlagInfoArr() As FlagInfo

Public MFP_TableIndex As New Dictionary

Public Function Init_RSCR()
    'TheHdw.StartStopwatch

    'soc rscr table read in
    Dim k As Integer

    Const start_row = 2
    Const end_col = 4
    
    If Flag_RSCR_INIT = False Then

        ' ==================== New Mbist RSCR ====================
        Dim ws As Worksheet
        Dim blockName As String
        Dim blockNum As Long

        blockNum = 0
        Set MbistBlockName = CreateObject("Scripting.Dictionary")
        MbistBlockName.compareMode = 1

        Dim sheetnames() As String

        sheetnames = TheExec.Job.GetSheetNamesOfType(DMGR_SHEET_TYPE_USER)

        Dim indx As Integer

        
        For indx = 0 To UBound(sheetnames)
             If sheetnames(indx) Like "RSCR*" Then
                blockName = Right(sheetnames(indx), Len(sheetnames(indx)) - 5)
                If Not MbistBlockName.Exists(blockName) Then
                    MbistBlockName.Add blockName, blockNum
                End If

                Dim MaxRow As Long
                MaxRow = Worksheets(sheetnames(indx)).UsedRange.Rows.Count
                MbistBlock(blockNum).MaxRow = MaxRow
                ReDim MbistBlock(blockNum).MemArray(MaxRow - 2)
                ReDim MbistBlock(blockNum).MemStrArray(MaxRow - 2)

                Dim arr1() As Variant
                Worksheets(sheetnames(indx)).Activate
                arr1 = Worksheets(sheetnames(indx)).range(Cells(start_row, 1), Cells(MaxRow, end_col)).Value
                Dim i As Integer

                For i = 1 To MaxRow - 1
                     MbistBlock(blockNum).MemArray(i - 1) = Int(arr1(i, 1))
                     MbistBlock(blockNum).MemStrArray(i - 1) = arr1(i, 2) & " " & arr1(i, 3) & " " & arr1(i, 4)
                Next i

                blockNum = blockNum + 1
            End If
        Next indx

        ' ==================== New Mbist RSCR ====================

        TheExec.Datalog.WriteComment "print: RSCR table initialized complete"

    End If
    Flag_RSCR_INIT = True

    'Debug.Print " 2new : " & TheHdw.ReadStopwatch

End Function

Public Function Mbist_RSCR(Shift_Pat As Pattern, MBIST_BLOCK As String, Optional Server As String)

On Error GoTo errHandler

    Dim SampleNum As Integer
    Dim CNumber_plus As Integer
    Dim testS As String
    Dim testS1 As String
    Dim full_str As String
    Dim BISTData(119) As Double
    Dim Mbist_repair_cycle As Long
    Dim capt As CaptType
    Dim numcap As New SiteLong
    Dim pre_trig As Long
    Dim PatData As New PinListData
    Dim PinData As New PinListData
    Dim PinPF As New PinListData
    Dim Failed_Pins() As String
    Dim maxDepth As Integer
    Dim HRAM_PFVar As Variant
    Dim HRAM_EXPECTVar As Variant
    Dim HRAM_DUTVar As Variant
    Dim RVal As New SiteDouble
    Dim file_name As String
    pre_trig = 0
    Dim k As Long
    Dim TestPatName As String, rtnPatternNames() As String, rtnPatternCount As Long
    Dim patt As Variant
    Dim kk  As Long
    Dim sne_str As String
    Dim patGup As String
    Dim mem_location As String, i As Long
    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    Dim site As Variant
    
    Dim M_maxrow As Long
    Dim M_MbistNum As Long
    Dim NewFmtAry() As String
    ReDim NewFmtAry(TheExec.sites.Existing.Count - 1)
    Dim OldFmtAry() As String
    Dim LineOffset As Long: LineOffset = 0
    Dim NumCapSum As Long: NumCapSum = 0
    Dim EmptyAry As Boolean: EmptyAry = True
    Dim RSCRHead As String: RSCRHead = "1"
    Dim patPassed As New SiteBoolean
    Dim TestNumber As Long
    Dim DecomPatName() As String
    
    SampleNum = 120
    CNumber_plus = 0 ' pattern has dummy cycle
    ''''''''''''''''''

    TheHdw.Digital.Patgen.HaltMode = tlHaltOnHRAMFull
    maxDepth = TheHdw.Digital.HRAM.maxDepth
    TheHdw.Digital.HRAM.Size = maxDepth
    TheHdw.Digital.HRAM.CaptureType = captFail
    TheHdw.Digital.HRAM.SetTrigger trigFail, False, 0

    rtnPatternNames = TheExec.DataManager.Raw.GetPatternsInSet(Shift_Pat.Value, rtnPatternCount)
    DecomPatName = Split(rtnPatternNames(0), ":")

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered 'SEC DRAM

    'RSCR site loop
    For Each site In TheExec.sites

        TheHdw.Patterns(Shift_Pat).Load
        TheHdw.Patterns(Shift_Pat).start ""
        TheHdw.Digital.Patgen.HaltWait
        patPassed = TheHdw.Digital.Patgen.PatternBurstPassed(site)

        numcap(site) = TheHdw.Digital.HRAM.CapturedCycles

        If numcap(site) = 0 Then
            NewFmtAry(site + LineOffset) = "RSCR2," & RSCRHead & "," & site & "," & Server & "," & DecomPatName(1) & ",NA,"
            If EmptyAry = True Then
                ReDim OldFmtAry(0)
                EmptyAry = False
                OldFmtAry(NumCapSum) = "Site " & site & "," & "Fail cycle at: NA" & ", Prime"
                NumCapSum = NumCapSum + 1
            Else
                ReDim Preserve OldFmtAry(UBound(OldFmtAry) + 1)
                OldFmtAry(NumCapSum) = "Site " & site & "," & "Fail cycle at: NA" & ", Prime"
                NumCapSum = NumCapSum + 1
            End If
        Else
            If EmptyAry = True Then
                ReDim OldFmtAry(numcap(site))
                EmptyAry = False
            Else
                ReDim Preserve OldFmtAry(UBound(OldFmtAry) + numcap(site))
            End If
            
            NewFmtAry(site + LineOffset) = "RSCR2," & RSCRHead & "," & site & "," & Server & "," & DecomPatName(1) & ","
            
            RVal(site) = TheHdw.Digital.HRAM.PatGenInfo(numcap(site) - 1, pgCycle)
            mem_location = "none"
            
            For i = 0 To numcap(site) - 1
                PinData = TheHdw.Digital.Pins("JTAG_TDO").HRAM.PinData(0, 1, numcap(site))

                '//MEMORY_CL52 cycle 2421 to cycle 3140
                '//MEMORY_CL51 cycle 3141 to cycle 7620
                '//MEMORY_CL27 cycle 8901 to cycle 9299
                '//MEMORY_CL26 cycle 10001 to cycle 10399
                '//MEMORY_CL17 cycle 11741 to cycle 12139
                '//MEMORY_CL16 cycle 12841 to cycle 13239
                'Mbist_repair_cycle
                Mbist_repair_cycle = TheHdw.Digital.HRAM.PatGenInfo(i, pgCycle)
                'Mbist_repair_cycle = Mbist_repair_cycle + 1    'start from cycle 1

                mem_location = "None"
                
                If Len(NewFmtAry(site + LineOffset)) < 250 Then
                    NewFmtAry(site + LineOffset) = NewFmtAry(site + LineOffset) & Mbist_repair_cycle & ","
                Else
                    LineOffset = LineOffset + 1
                    'ReDim Preserve NewFmtAry((TheExec.sites.Existing.Value) + LineOffset)
                    ReDim Preserve NewFmtAry((TheExec.sites.Existing.Count) - 1 + LineOffset)
                    NewFmtAry(site + LineOffset) = "RSCR2," & RSCRHead & "," & site & "," & Server & "," & DecomPatName(1) & "," & Mbist_repair_cycle & ","
                End If

                If MbistBlockName.Exists(MBIST_BLOCK) Then
                    M_MbistNum = MbistBlockName.Item(MBIST_BLOCK)
                    M_maxrow = MbistBlock(M_MbistNum).MaxRow - 2

                    For k = 0 To M_maxrow
                        If Mbist_repair_cycle = MbistBlock(M_MbistNum).MemArray(k) Then mem_location = MbistBlock(M_MbistNum).MemStrArray(k)
                    Next k
                Else
                    mem_location = "Block-non-define"
                End If
                
                OldFmtAry(NumCapSum) = "Site " & site & "," & "Fail cycle at: " & Mbist_repair_cycle & ",Mem:" & mem_location
                NumCapSum = NumCapSum + 1
            Next i
        End If

        '///print pattern result 170626
        TestNumber = TheExec.sites.Item(site).TestNumber

        If patPassed Then
            Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
        Else
            Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
        End If
        '///print pattern result 170626
    Next site

    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
    For i = 0 To UBound(NewFmtAry)
        If NewFmtAry(i) <> "" Then
            NewFmtAry(i) = Left(NewFmtAry(i), Len(NewFmtAry(i)) - 1)
            TheExec.Datalog.WriteComment NewFmtAry(i)
        End If
    Next i
    '===================================================================================
    If TheExec.Flow.EnableWord("RSCR_MP") <> True Then

            TheExec.Datalog.WriteComment "Mbist repair information shift start"
            
            For i = 0 To UBound(OldFmtAry)
                TheExec.Datalog.WriteComment OldFmtAry(i)
            Next i
        
            TheExec.Datalog.WriteComment "Mbist repair information shift end"
    End If
    '===================================================================================
            DebugPrintFunc Shift_Pat.Value
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment ("Site " & site & "," & "RSCR Err.")
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function TurnOnEfusePwrPins_Mbist(FusePower As String)

    'Escalate VDD18_EFUSE0 and VDD18_EFUSE1 according to Fiji
    'test plan (slower than 1.8v/30us)
    
    DCVS_PowerOn_I_Meter FusePower, 1.8, 0.2, 0.001, 0.002, 10, 0.018   'use 18 ms to power up

End Function


Public Function TurnOffEfusePwrPins_Mbist(FusePower As String)

    'Decline VDD18_EFUSE0 and VDD18_EFUSE1 according to Fiji
    'test plan (slower than 1.8v/30us)
    
    Dim CurrentVoltage As Double
    
    CurrentVoltage = TheHdw.DCVS.Pins(FusePower).Voltage.Main.Value
    DCVS_PowerOff_I_Meter FusePower, CurrentVoltage, 0.2, 0.001, 0.002, 10, 0.018   'use 18 ms to power down


End Function
Public Function MbistRetentionLevelWait_and_lowDown_power(mS_Time As Double, Pwr_pins As PinList, Low_Vol As Double)
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered 'SEC DRAM
    Dim original_vol As Double

    original_vol = TheHdw.DCVS.Pins(Pwr_pins).Voltage.Main.Value
    TheHdw.DCVS.Pins(Pwr_pins).Voltage.Main.Value = Low_Vol

    TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Datalog.WriteComment "*print: Lower pin: " & Pwr_pins & " to " & Low_Vol & " *"
    TheExec.Datalog.WriteComment "*************************************************"

    DebugPrintFunc ""
    TheHdw.Wait mS_Time * 0.001
    
    
    TheHdw.DCVS.Pins(Pwr_pins).Voltage.Main.Value = original_vol

    TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
    TheExec.Datalog.WriteComment "*************************************************"


End Function

Public Function MbistRetentionLevelWait(mS_Time As Double, Retention_Voltage As Double, Retention_Pins As PinList, RampStep As Double, Optional RampWaitTime As Double = 0, Optional WaitTimeOnly As Boolean = False)
   ' thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered  'SEC DRAM
    
    ''SWLINZA20171103, for ramp up/down retention voltage

    Dim Retention_Pins_Ary() As String
    Dim Retention_Pins_count As Long
    Dim RampDown_Time As Double: RampDown_Time = RampWaitTime 'RampDown_Time = 0
    Dim RampDown_Step As Double
    Dim Original_voltage() As Double
    Dim DropVoltage() As Double
    Dim DropVoltage_perStep() As Double
    Dim Voltage_from_HW As String
    Dim i, j As Integer
    
    If WaitTimeOnly = False Then
        If RampStep = 0 Then
            RampDown_Step = 20 ' default RampDown_Step = 20
        Else
            RampDown_Step = RampStep
        End If
       
        TheExec.DataManager.DecomposePinList Retention_Pins, Retention_Pins_Ary(), Retention_Pins_count
        ReDim Original_voltage(Retention_Pins_count - 1) As Double
        ReDim DropVoltage(Retention_Pins_count - 1) As Double
        ReDim DropVoltage_perStep(Retention_Pins_count - 1) As Double
        
        For i = 0 To Retention_Pins_count - 1
            Original_voltage(i) = FormatNumber(TheHdw.DCVS.Pins(Retention_Pins_Ary(i)).Voltage.Value, 3)
            DropVoltage(i) = Original_voltage(i) - Retention_Voltage
            DropVoltage_perStep(i) = FormatNumber((DropVoltage(i) / RampDown_Step), 3)
        Next i
    
        '--------- Ramp down for retention voltage ------'
        For i = 0 To RampDown_Step - 1
            For j = 0 To Retention_Pins_count - 1
                If i = RampDown_Step - 1 Then
                    TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Value = Retention_Voltage
                Else
                    TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Value = Original_voltage(j) - DropVoltage_perStep(j) * i
                End If
            Next j
            TheHdw.Wait RampDown_Time / RampDown_Step
        Next i
        
        Voltage_from_HW = ""
        '--------- Read back retention voltage from HW ------'
        For j = 0 To Retention_Pins_count - 1
            If j = 0 Then
                Voltage_from_HW = CStr(FormatNumber(TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Value, 3)) & " V"
            Else
                Voltage_from_HW = Voltage_from_HW & "," & CStr(FormatNumber(TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Value, 3)) & " V"
            End If
        Next j
            
        '----- Retention Wait time 100 ms ------
        TheHdw.Wait mS_Time * 0.001
    
        TheExec.Flow.TestLimit mS_Time, PinName:="Wait_Time", Unit:=unitCustom, customUnit:="mSec"
        TheExec.Datalog.WriteComment "*************************************************"
        TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
        TheExec.Datalog.WriteComment "*print: MbistRetention Pins " & Retention_Pins
        TheExec.Datalog.WriteComment "*print: MbistRetention Volt " & Voltage_from_HW
        TheExec.Datalog.WriteComment "*************************************************"
        DebugPrintFunc ""
    ''
        '--------- Ramp up for retention voltage ------'
        For i = 0 To RampDown_Step - 1
            For j = 0 To Retention_Pins_count - 1
                If i = RampDown_Step - 1 Then
                    TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Value = Original_voltage(j)
                Else
                    TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Value = Retention_Voltage + DropVoltage_perStep(j) * i
                End If
            Next j
            TheHdw.Wait RampDown_Time / RampDown_Step
        Next i
    Else
        '----- Retention Wait time 100 ms ------
        TheHdw.Wait mS_Time * 0.001
        TheExec.Flow.TestLimit mS_Time, PinName:="Wait_Time", Unit:=unitCustom, customUnit:="mSec"
        TheExec.Datalog.WriteComment "*************************************************"
        TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
        TheExec.Datalog.WriteComment "*************************************************"
        DebugPrintFunc ""
    End If
    
End Function

Public Function Init_MBISTFailBlock() 'Multi Sheets

    'TheHdw.StartStopwatch

    Dim MBISTFAILBLOCKSHEETS(0) As String
    
    MBISTFAILBLOCKSHEETS(0) = "MBISTFailBlock"
    
'    MBISTFAILBLOCKSHEETS(0) = "MBISTFailBlock_CPU"
'    MBISTFAILBLOCKSHEETS(1) = "MBISTFailBlock_GFX"
'    MBISTFAILBLOCKSHEETS(2) = "MBISTFailBlock_SOC"
    
    
    Dim k As Long
    Dim M As Long
    Dim n As Long
    Dim i As Long
        
    M = 0
    n = 0
    Dim sheetIdx As Integer
    sheetIdx = 0
    Const start_row = 2
    Const end_col = 5
    
    If Flag_MBISTFailBlock_INIT = False Then
        ReDim tpCycleBlockInfor(300)
        For sheetIdx = 0 To UBound(MBISTFAILBLOCKSHEETS)
            
            Dim MaxRow As Long
            
            Dim arr1() As Variant
        
            MaxRow = Worksheets(MBISTFAILBLOCKSHEETS(sheetIdx)).UsedRange.Rows.Count
            
            Dim maxcolumn As Long
            maxcolumn = Worksheets(MBISTFAILBLOCKSHEETS(sheetIdx)).UsedRange.Columns.Count
            
            Worksheets(MBISTFAILBLOCKSHEETS(sheetIdx)).Activate
            
            arr1 = Worksheets(MBISTFAILBLOCKSHEETS(sheetIdx)).range(Cells(start_row, 1), Cells(MaxRow, maxcolumn)).Value
            

            
            For k = 1 To MaxRow - 1
                
               ' If arr1(2, 1) = "" Then Exit For
                If (k = 1) And (sheetIdx = 0) Then
                    ReDim tpCycleBlockInfor(M).tpMbistCycleBlock(1800)
                    ReDim tpCycleBlockInfor(M).strDecsName(maxcolumn - 6)
                    ReDim tpCycleBlockInfor(M).strServerName(maxcolumn - 6)
                    tpCycleBlockInfor(M).strPattName = arr1(k, 1)
                Else
                    If tpCycleBlockInfor(M).strPattName = arr1(k, 1) Then
                        If n > 1800 Then
                            ReDim Preserve tpCycleBlockInfor(M).tpMbistCycleBlock(n + 200)
                        End If
                        
                    Else
                        ReDim Preserve tpCycleBlockInfor(M).tpMbistCycleBlock(n - 1)
                        
                        n = 0
                        M = M + 1
                        
                        If M > 300 Then
                            ReDim Preserve tpCycleBlockInfor(M + 20)
                        End If
                        
                        ReDim tpCycleBlockInfor(M).tpMbistCycleBlock(1800)
                        ReDim tpCycleBlockInfor(M).strDecsName(maxcolumn - 6)
                        ReDim tpCycleBlockInfor(M).strServerName(maxcolumn - 6)
                        
                        tpCycleBlockInfor(M).strPattName = arr1(k, 1)
                    End If
                End If
                
                tpCycleBlockInfor(M).tpMbistCycleBlock(n).strBlaclName = arr1(k, 2)
                tpCycleBlockInfor(M).tpMbistCycleBlock(n).lVector = Int(arr1(k, 3))
                tpCycleBlockInfor(M).tpMbistCycleBlock(n).lCycle = Int(arr1(k, 4))
                tpCycleBlockInfor(M).tpMbistCycleBlock(n).strCompare = arr1(k, 5)

                For i = 6 To maxcolumn
                    Dim tempAry() As String
                    If i = 6 Then
                        tempAry() = Split(",", ",")
                    ElseIf arr1(k, i) = Empty Then
                        tempAry() = Split(",", ",")
                    Else
                        tempAry() = Split(arr1(k, i), ",")
                        If UBound(tempAry()) = 0 Then
                            ReDim Preserve tempAry(1)
                            tempAry(1) = ""
                        End If
                    End If
                    tpCycleBlockInfor(M).strServerName(i - 6) = tempAry(0)
                    tpCycleBlockInfor(M).strDecsName(i - 6) = tempAry(1)
                Next i
                    
                n = n + 1
            Next k
    
        Next sheetIdx
        
        ReDim Preserve tpCycleBlockInfor(M).tpMbistCycleBlock(n - 1)
        ReDim Preserve tpCycleBlockInfor(M)
        'TheExec.Datalog.WriteComment RepeatChr("*", 120)
        TheExec.Datalog.WriteComment "print: MBISTFailBlock table initialized complete"
    End If
    Flag_MBISTFailBlock_INIT = True
    
    'Debug.Print " Init_MBISTFailBlock new : " & TheHdw.ReadStopwatch
        
End Function

''' add 20160629  webster
Public Function GetFlagInfoArrIndex(flagName As String) As Long
On Error GoTo errHandler
Dim funcName As String: funcName = "MbistFP_FlagPrintTest"
Dim lIdxTemp As Long

GetFlagInfoArrIndex = -1

For lIdxTemp = 0 To UBound(tyFlagInfoArr)
    If tyFlagInfoArr(lIdxTemp).flagName = flagName Then
        GetFlagInfoArrIndex = lIdxTemp
        Exit For
    End If
Next lIdxTemp

If GetFlagInfoArrIndex = -1 Then
    TheExec.Datalog.WriteComment "<Warnning> the flag(" & flagName & ") can not be found in MBISTFailBlock excel sheet"
End If

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MbistRetentionWait(mS_Time As Double)
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered 'SEC DRAM

    TheHdw.Wait mS_Time * 0.001

    TheExec.Flow.TestLimit mS_Time, PinName:="Wait_Time", Unit:=unitCustom, customUnit:="mSec"
    TheExec.Datalog.WriteComment "*************************************************"
    TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
    TheExec.Datalog.WriteComment "*************************************************"
    DebugPrintFunc ""

End Function

