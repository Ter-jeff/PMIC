Attribute VB_Name = "VBT_Write_SPIROM"
Option Explicit

Global Uniform_256KB_Sector As Boolean
Global Total_Count As Long
Const Mem32MByte As Long = 268435456
Const Mem16MByte As Long = 134217728
Const Addr32MByte As Long = 33554432
Const Addr16MByte As Long = 16777216
Const Power = 3

Dim SPI_Binary_Signal As New DSPWave
Dim SPI_Binary_32MegErase As New DSPWave
Dim SPI_Binary_32MegCode As New DSPWave
Dim SPI_Binary_32MegCode_PP As New DSPWave

Public SPI_Flag_0 As New SiteBoolean
Public SPI_Flag_1 As New SiteBoolean
Public SPI_Flag_2 As New SiteBoolean
Public SPI_Flag_3 As New SiteBoolean
Public SPI_Flag_4 As New SiteBoolean
Public SPI_Flag_5 As New SiteBoolean
Public SPI_Flag_6 As New SiteBoolean
Public SPI_Flag_7 As New SiteBoolean

Public SPI_Flag_temp0 As New SiteLong
Public SPI_Flag_temp1 As New SiteLong
Public SPI_Flag_temp2 As New SiteLong
Public SPI_Flag_temp3 As New SiteLong
Public SPI_Flag_temp4 As New SiteLong
Public SPI_Flag_temp5 As New SiteLong
Public SPI_Flag_temp6 As New SiteLong
Public SPI_Flag_temp7 As New SiteLong
Public CheckSumResult16M As Long
Public CheckSumResult32M As Long
Public SPI_Flag_Sum_16M As New SiteLong
Public SPI_Flag_Sum_32M As New SiteLong
'Public Site As Variant 'remove from Public
Public SpiromCodeFile As String

Public Sub SPI_ROM_Written_record()

Dim site As Variant

Set SPI_Binary_Signal = Nothing
Set SPI_Binary_32MegErase = Nothing
Set SPI_Binary_32MegCode = Nothing
Set SPI_Binary_32MegCode_PP = Nothing

For Each site In TheExec.sites.Active
    If write_spirom = True Then
        TheExec.Datalog.WriteComment ("site(" & site & ") has been writen")
    End If
    write_spirom = False
Next site

End Sub



Public Function SPIROM_32M_Check(PatName As Pattern, RelayOnPins As PinList)

Dim site As Variant
Dim check_result As New SiteBoolean
Dim result_flag As New SiteLong
Dim result_flag_temp As New SiteLong
Dim check_32M_flag As New SiteLong
Dim check_16M_flag As New SiteLong
Dim error_flag As New SiteLong
Dim RomSize As New SiteLong
Dim RomFileName1 As String
Dim AndResult As New SiteLong
Dim SPIRunCheckCase As New SiteVariant

    On Error GoTo ErrorHandler
    
          check_32M_flag = 0
          check_16M_flag = 0
          error_flag = 0
          
          Dim CurrPath As String
          CurrPath = Application.ActiveWorkbook.Path
          
          If TheExec.Flow.EnableWord("SPIROM_1_Write") = True Then
                 RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).Value
                 CurrPath = CurrPath & Mid(RomFileName1, 2) '20160322
                 'If Dir(CurrPath) = Empty Or RomFileName1 = "" Then  it will cause the error at OSAT 20160513
                
                 If RomFileName1 = "" Then
                        For Each site In TheExec.sites
                            TheExec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                            TheExec.Datalog.WriteComment ("Romcode File Did NOT Specify. Plese Check SpiromCodeFile Sheet  ")
                        Next site
                        Exit Function ' 20160322
                 End If
                 
                 '' check if the file exists or not 20160513
                  If Dir(RomFileName1) = "" Then
                        For Each site In TheExec.sites
                            TheExec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                            TheExec.Datalog.WriteComment (" Binary File Does NOT Exist. Please Check Binary Directory  ")
                        Next site
                        Exit Function ' 20160513
                 End If
                 For Each site In TheExec.sites
                        RomSize = SPIROM_Get_RomSize_VBT(RomFileName1)
                 Next site
          End If
          
          If TheExec.Flow.EnableWord("SPIROM_2_Write") = True Then
                RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).Value
                CurrPath = CurrPath & Mid(RomFileName1, 2)  '20160322
                 'If Dir(CurrPath) = Empty Or RomFileName1 = "" Then  it will cause the error at OSAT 20160513
                
                 If RomFileName1 = "" Then
                        For Each site In TheExec.sites
                            TheExec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                            TheExec.Datalog.WriteComment ("Romcode File Did NOT Specify. Plese Check SpiromCodeFile Sheet  ")
                        Next site
                        Exit Function ' 20160322
                 End If
                 
                 '' check if the file exists or not 20160513
                  If Dir(RomFileName1) = "" Then
                        For Each site In TheExec.sites
                            TheExec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                            TheExec.Datalog.WriteComment (" Binary File Does NOT Exist. Please Check Binary Directory  ")
                        Next site
                        Exit Function ' 20160513
                 End If
                 For Each site In TheExec.sites
                        RomSize = SPIROM_Get_RomSize_VBT(RomFileName1)
                 Next site
            End If
                   
          If TheExec.Flow.EnableWord("SPIROM_3_Write") = True Then
                RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).Value
                CurrPath = CurrPath & Mid(RomFileName1, 2)  '20160322
                 'If Dir(CurrPath) = Empty Or RomFileName1 = "" Then  it will cause the error at OSAT 20160513
                
                 If RomFileName1 = "" Then
                        For Each site In TheExec.sites
                            TheExec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                            TheExec.Datalog.WriteComment ("Romcode File Did NOT Specify. Plese Check SpiromCodeFile Sheet  ")
                        Next site
                        Exit Function ' 20160322
                 End If
                 
                 '' check if the file exists or not 20160513
                  If Dir(RomFileName1) = "" Then
                        For Each site In TheExec.sites
                            TheExec.Flow.TestLimit 0, 1, 1, , , , , , "RomFileError"
                            TheExec.Datalog.WriteComment (" Binary File Does NOT Exist. Please Check Binary Directory  ")
                        Next site
                        Exit Function ' 20160513
                 End If
                 For Each site In TheExec.sites
                        RomSize = SPIROM_Get_RomSize_VBT(RomFileName1)
                 Next site
            End If
                   
           TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
          ' Level & Timing
'''          TheHdw.digital.ApplyLevelsTiming True, True, True, tlPowered
'''          TheHdw.wait 0.5
'''          TheHdw.Patterns(PatName).Load
'''          TheHdw.Patterns(PatName).Start ""
'''          TheHdw.digital.Patgen.HaltWait
          ' modify by Chihome 140418
           TheHdw.Digital.ApplyLevelsTiming True, True, True, tlUnpowered, "SPI1_SSIN", "SPI1_SCLK,SPI1_MOSI", "SPI1_MISO"
'          TheHdw.wait 1
          TheHdw.Patterns(PatName).Load

          TheHdw.Patterns(PatName).start ""
          TheHdw.Digital.Patgen.HaltWait
          For Each site In TheExec.sites
''            If thehdw.Digital.Patgen.PatternBurstPassedPerSite(site) = False Then
            If TheHdw.Digital.TimeDomains("SPIROM_Tdomain").Patgen.PatternBurstPassedPerSite = False Then
                TheHdw.Wait 2 '10
                TheHdw.Patterns(PatName).start ""
                TheHdw.Digital.Patgen.HaltWait
            End If
          Next site
'          TheHdw.wait 10
'          TheHdw.Patterns(PatName).Start ""
'          TheHdw.digital.Patgen.HaltWait
          
         For Each site In TheExec.sites
'''              check_result = thehdw.Digital.Patgen.PatternBurstPassedPerSite(site)
              check_result = TheHdw.Digital.TimeDomains("SPIROM_Tdomain").Patgen.PatternBurstPassedPerSite
              AndResult = 15
              If check_result = True Then
                   result_flag = 1 '32M
                   result_flag_temp = 1
               Else
                   result_flag = 0 '16M
                   result_flag_temp = 2
               End If
         Next site
          
         For Each site In TheExec.sites
             AndResult = result_flag_temp(site) And 15 And AndResult
             If AndResult = 1 Then SPIRunCheckCase = "SPIROMTesting"
             If RomSize > 16777216 * 2 Then SPIRunCheckCase = "ErrorTestingB"
             If AndResult = 0 Then SPIRunCheckCase = "ErrorTestingA"  ' check whether all spirom devices are the same or not!
             If AndResult = 2 And RomSize > 16777216 Then SPIRunCheckCase = "ErrorTestingC"  ' 16M testing!! all spirom devices are 16M, but romsize is large than 32M
             If AndResult = 2 Then SPIRunCheckCase = "ErrorTestingD"  ' either one of the spirom =16M or simply the functional test fails such as relay broken 20160415
         Next site
         
         
         For Each site In TheExec.sites
                Select Case True
                
                    Case SPIRunCheckCase = "SPIROMTesting"
SPIROMTesting: 'no any error on SPIROM check!
                                If AndResult = 1 And RomSize > 16777216 Then
                                    check_32M_flag = 1
                                    TheExec.Flow.TestLimit check_32M_flag, 1, 1, , , , , , "SPIROM_32M_Testing"
                                    TheExec.Datalog.WriteComment ("The Rom File of site(" & CStr(site) & ")is " & RomSize)
                                    TheExec.Datalog.WriteComment ("The RomSize of site(" & CStr(site) & ")is " & "32MB") '20160328
                                ElseIf AndResult = 1 And RomSize < 16777216 Then
                                    check_32M_flag = 1
                                    TheExec.Flow.TestLimit check_32M_flag, 1, 1, , , , , , "SPIROM_32M_Testing"
                                    TheExec.Datalog.WriteComment ("The Rom File of site(" & CStr(site) & ")is " & RomSize)
                                    TheExec.Datalog.WriteComment ("The RomSize of site(" & CStr(site) & ")is " & "32MB") '20160328
                                End If
                                
                                
                    Case SPIRunCheckCase = "ErrorTestingA"
ErrorTestingA: ' check whether all spirom devices are the same or not!
                                TheExec.Datalog.WriteComment ("Please Makse Sure All SPIROM Devices In the Daughter Are The Same!!")
                                If result_flag = 1 Then
                                    TheExec.Datalog.WriteComment ("The Site(" & CStr(site) & ") is 32M SPIROM Device.")
                                Else
                                    TheExec.Datalog.WriteComment ("The Site(" & CStr(site) & ") is 16M SPIROM Device.")
                                End If
                                TheExec.Flow.TestLimit AndResult, 1, 1, , , , , , "Error_Testing"
                                MsgBox ("Please make sure SPIROM devices are all the same")
                    

                    Case SPIRunCheckCase = "ErrorTestingB"
ErrorTestingB: ' check whether rom file size > 32M
                                TheExec.Datalog.WriteComment ("Please Check the Rom File Size")
                                TheExec.Flow.TestLimit RomSize, 1, 1, , , , , , "Error_Testing"


                    Case SPIRunCheckCase = "ErrorTestingC"
ErrorTestingC: ' all spirom devices are 16M, but romsize is large than 32M
                                TheExec.Flow.TestLimit error_flag, 1, 1, , , , , , "Error_Testing"
                                TheExec.Datalog.WriteComment ("The Site(" & CStr(site) & ") is 16M SPIROM Device.")
                                TheExec.Datalog.WriteComment ("RomFile of Site(" & CStr(site) & ")=> " & RomFileName1 & " is correct?!")
                                MsgBox ("Please check ROM file !!")


                    Case SPIRunCheckCase = "ErrorTestingD"
ErrorTestingD: ' either one of the spirom =16M or simply the functional test fails such as relay broken 20160415
                                TheExec.Flow.TestLimit AndResult, 1, 1, , , , , , "SPIROM_32M_Fail"
                                
            End Select
        Next site
                   
Exit Function

ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function







Public Function SPIROM_Continuity(RelayOnPins As PinList) As Long
    On Error GoTo errHandler
    
    Dim PPMUMeasure As New PinListData
        
'''    TheHdw.Utility.pins("K80_SPI0_SCLK, K84_SPI0_MISO").State = tlUtilBitOn
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
    TheHdw.Digital.Pins("SPIROM_PINS").Disconnect
                       
    TheHdw.Wait 0.05
               
    With TheHdw.PPMU.Pins("SPIROM_PINS")
         .Connect
         .ForceI -0.0002, -2 * mA, -0.2, -1
         .Gate = tlOn
    End With
        
    TheHdw.Wait 0.05

    TheHdw.PPMU.Pins("SPIROM_PINS").Test
    
    TheHdw.Wait 0.05
    
    With TheHdw.PPMU.Pins("SPIROM_PINS")
        .ForceI 0
        .Gate = tlOff
        .Disconnect
    End With
    
    TheHdw.Digital.Pins("SPIROM_PINS").Connect
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function




Public Function PA_WriteEnable(argc As Long, argv() As String) As Long
  
    TheHdw.Wait 0.5
    TheHdw.Patterns(".\PATTERN\SPIROM\SPIROMWriteEnableNCheckWEL.PAT").ValidateThreading
    TheHdw.Wait 0.5
End Function



Public Function PA_Erase(argc As Long, argv() As String) As Long
  
    TheHdw.Wait 0.5
'    TheHdw.patterns(".\PATTERN\SPIROMBulkEraseNCheckWIP.PAT").ValidateThreading
    TheHdw.Patterns(".\PATTERN\SPIROM\SPIROMBulkEraseNCheckWIP.PAT").Threading.Enable = True
    TheHdw.Wait 0.5
End Function



Public Function BE_TimeOut(argc As Long, argv() As String) As Long
    
    TheHdw.Digital.Patgen.TimeOut = 770
    
    TheHdw.Wait 1
    
End Function



Public Function SPIROM_FL_PP_VBT(PatName As Pattern, RomFileName1 As String, RelayOnPins As PinList) As Long
    
    On Error GoTo ErrorHandler
    Dim Rom_Data As Byte
    Dim DspArray(258) As Long
    Dim TotalDspArray() As Long
    Dim Count As Long
    Dim DspSrcWave As New DSPWave
    Dim DspSrcWave1 As New DSPWave
    Dim DspRefWave As New DSPWave
    Dim DspRefWave1 As New DSPWave
    Dim SrcWaveName As String
    Dim SPI_DSSCCap_Signal As New DSPWave
    Dim result As New SiteDouble
    Dim WaveCount As Long
   
    PatName = Worksheets("Pattern_Sets_SPIROM").Cells(8, 5).Value
 
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
    
    
    Const SrcDataSize = 259
    Total_Count = 0
    WaveCount = 0
    Close #1
    
    ' Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Patterns(PatName).Load
    TheHdw.Digital.Patgen.TimeOut = 10
              
    Open RomFileName1 For Binary As #1
      
   
'=============================================================

    While (Not EOF(1)) And (Total_Count < 16777216) 'bigrom size 4936944
        For Count = 3 To SrcDataSize - 1
            DspArray(Count) = 2 ^ 8 - 1
        Next Count
        
        Count = 3
        
        While ((Not EOF(1)) And (Count < SrcDataSize))
            Get #1, Total_Count + 1, Rom_Data
            DspArray(Count) = Rom_Data
            Count = Count + 1
            Total_Count = Total_Count + 1
            'Debug.Print Rom_Data
        Wend
        
        'Address
        Dim Address(23) As String
        Dim Temp_data As String
        Dim DecToBin As String
        Dim i As Integer
        Dim tempA As New SiteLong
        Dim tempB As New SiteLong
        Dim tempC As New SiteLong
        Dim temp0_7 As New SiteLong
        Dim temp8_15 As New SiteLong
        Dim temp16_23 As New SiteLong
        Dim site As Variant
        Dim TempArray(23) As Long
        
        DecToBin = vbNullString
        Count = 0
        
        If WaveCount = 0 Then
                For i = 0 To 23
                    Address(i) = 0
                    TempArray(Count) = Address(i)
                Next i
            Else
                For i = 0 To 23
                    DecToBin = CStr(((WaveCount * 256) And 2 ^ i) / 2 ^ i) & DecToBin
                    Address(23 - i) = Left(DecToBin, 1)
                    
                    TempArray(i) = Address(23 - i)
                Next i
                
                temp0_7(site) = 0
                temp8_15(site) = 0
                temp16_23(site) = 0
                
                For i = 0 To 7
                    tempA(site) = TempArray(i) * 2 ^ (i)
                    temp0_7(site) = temp0_7(site) + tempA(site)
                Next i
               
                For i = 8 To 15
                    tempB(site) = TempArray(i) * 2 ^ (i - 8)
                    temp8_15(site) = temp8_15(site) + tempB(site)
                Next i
                For i = 16 To 23
                    tempC(site) = TempArray(i) * 2 ^ (i - 16)
                    temp16_23(site) = temp16_23(site) + tempC(site)
                Next i

                DspArray(0) = temp16_23(site)
                DspArray(1) = temp8_15(site)
                DspArray(2) = temp0_7(site)
                
        End If
                       

        'WaveDefinitions
        'SrcWaveName = "Src_Wave_" & WaveCount
        DspSrcWave.Data = DspArray
        Call TheExec.WaveDefinitions.CreateWaveDefinition("SrcData", DspSrcWave, True)
        DspRefWave = DspSrcWave
        'DspSrcWave.Plot

        'Insert DSSC Loading & Pattern Loading & Pattern Execution.
        With TheHdw.DSSC.Pins("SPI1_MOSI").Pattern(PatName).Source
            .Signals.Add "DSSC_Src"
            .Signals.DefaultSignal = "DSSC_Src"
            .Signals("DSSC_Src").WaveDefinitionName = "SrcData"
'''            .Signals.Item("DSSC_Src").SampleSize = 256
            .Signals.Item("DSSC_Src").SampleSize = 259
            .Signals("DSSC_Src").Amplitude = 1
            If glb_TesterType = "Jaguar" Then .Signals("DSSC_Src").LoadSettings
        End With

'''        TheHdw.Wait 0.05

        TheHdw.Patterns(PatName).Test pfNever, 0

        TheHdw.Digital.Patgen.HaltWait
        
'''=============================================================
''        Dim DSSCCap_All_Signal_2 As New DSPWave
''        TheHdw.Wait 0.5
''        TheHdw.Digital.patterns.Pat(".\PATTERN\SPIROMDedug.PAT").Load
''        TheHdw.Digital.Patgen.TimeOut = 800
''         With TheHdw.DSSC.Pins("SPI0_MISO").Pattern(".\PATTERN\SPIROMDedug.PAT").Capture
''            .Signals.Add "DSSC_Cap_All_2"
''            .Signals("DSSC_Cap_All_2").Offset = 0
''            .Signals.Item("DSSC_Cap_All_2").SampleSize = 256
''            .Signals("DSSC_Cap_All_2").LoadSettings
''        End With
''        TheHdw.Wait 0.05
''        TheHdw.patterns(".\PATTERN\SPIROMDedug.PAT").test pfNever, 0
''        TheHdw.Digital.Patgen.HaltWait
''        TheHdw.Wait 0.05
''        DSSCCap_All_Signal_2 = TheHdw.DSSC.Pins("SPI0_MISO").Pattern(".\PATTERN\SPIROMDedug.PAT").Capture.Signals("DSSC_Cap_All_2").DSPWave
'''=============================================================

        WaveCount = WaveCount + 1
''''        Debug.Print WaveCount
                   
    Wend
'======================================================== =====
    
    Close #1
    
    Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function





Public Function SPIROM_Get_RomSize_VBT(RomFileName1 As String) As Long
    
    On Error GoTo ErrorHandler
    Dim Rom_Data As Byte
    Dim DspArray(258) As Long
    Dim TotalDspArray() As Long
    Dim Count As Long
    Dim SPI_DSSCCap_Signal As New DSPWave
    Dim result As New SiteDouble
    Dim WaveCount As Long
    Dim RomSize As Long
   
'''    TheHdw.Utility.Pins("K3_SPI0_SCLK_SSIN_ROM,K5_SPI0_MISO_MOSI_ROM").State = tlUtilBitOn
    
    
    Const SrcDataSize = 259
    Total_Count = 0
    WaveCount = 0
    Close #1

    Open RomFileName1 For Binary As #1
      
'=============================================================

    While (Not EOF(1)) And (Total_Count < 33554432)     '256MBits
        For Count = 3 To SrcDataSize - 1
            DspArray(Count) = 2 ^ 8 - 1
        Next Count
        
        Count = 3
        
        While ((Not EOF(1)) And (Count < SrcDataSize))
            Get #1, Total_Count + 1, Rom_Data
            DspArray(Count) = Rom_Data
            Count = Count + 1
            Total_Count = Total_Count + 1
            'Debug.Print Rom_Data
        Wend
        WaveCount = WaveCount + 1
    Wend
'======================================================== =====
    
    Close #1
    SPIROM_Get_RomSize_VBT = Total_Count
    
    Exit Function
    
ErrorHandler:

        If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function SPIROM_FL_PP_VBT_32MByte(PatName As Pattern, RomFileName1 As String, RelayOnPins As PinList) As Long
    
    On Error GoTo ErrorHandler
    Dim Rom_Data As Byte
    Dim DspArray(259) As Long
    Dim TotalDspArray() As Long
    Dim Count As Long
    Dim DspSrcWave As New DSPWave
    Dim DspSrcWave1 As New DSPWave
    Dim DspRefWave As New DSPWave
    Dim DspRefWave1 As New DSPWave
    Dim SrcWaveName As String
    Dim SPI_DSSCCap_Signal As New DSPWave
    Dim result As New SiteDouble
    Dim WaveCount As Long
    Dim PatCount As Long
    Dim PattArray() As String
    Dim patt As String
    
    Const SrcDataSize = 260
    Total_Count = 0
    WaveCount = 0
    Close #1
    
''    PatName = Worksheets("PatSets_SPIROM").Cells(11, 5).Value
  
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
   
    ' Level & Timing

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call PATT_GetPatListFromPatternSet(PatName.Value, PattArray, PatCount)
    patt = CStr(PattArray(0))
    TheHdw.Patterns(patt).Load
    TheHdw.Digital.Patgen.TimeOut = 10
    
    'Address
    Dim Address(31) As String
    Dim Temp_data As String
    Dim DecToBin As String
    Dim i As Integer
    Dim tempA As New SiteLong
    Dim tempB As New SiteLong
    Dim tempC As New SiteLong
    Dim tempD As New SiteLong
    Dim temp0_7 As New SiteLong
    Dim temp8_15 As New SiteLong
    Dim temp16_23 As New SiteLong
    Dim temp24_31 As New SiteLong
    Dim site As Variant
    Dim TempArray(31) As Long
        
    
    Open RomFileName1 For Binary As #1
      
    '=============================================================
    
    While (Not EOF(1)) And (Total_Count < Addr32MByte)
        For Count = 4 To SrcDataSize - 1
            DspArray(Count) = 2 ^ 8 - 1
        Next Count
        
        Count = 4
        
        While ((Not EOF(1)) And (Count < SrcDataSize))
            Get #1, Total_Count + 1, Rom_Data
            DspArray(Count) = Rom_Data
            Count = Count + 1
            Total_Count = Total_Count + 1
            'Debug.Print Rom_Data
        Wend
        
        
        DecToBin = vbNullString
        Count = 0
        
        If WaveCount = 0 Then
            For i = 0 To 31
                Address(i) = 0
                TempArray(Count) = Address(i)
            Next i
        Else
            For i = 0 To 31
                If i < 25 Then
                    DecToBin = CStr(((WaveCount * 256) And 2 ^ i) / 2 ^ i) & DecToBin
                Else
                    DecToBin = "0" & DecToBin
                End If
                Address(31 - i) = Left(DecToBin, 1)
                TempArray(i) = Address(31 - i)
            Next i
            
            temp0_7(site) = 0
            temp8_15(site) = 0
            temp16_23(site) = 0
            temp24_31(site) = 0
            
            For i = 0 To 7
                tempA(site) = TempArray(i) * 2 ^ (i)
                temp0_7(site) = temp0_7(site) + tempA(site)
            Next i
           
            For i = 8 To 15
                tempB(site) = TempArray(i) * 2 ^ (i - 8)
                temp8_15(site) = temp8_15(site) + tempB(site)
            Next i
            
            For i = 16 To 23
                tempC(site) = TempArray(i) * 2 ^ (i - 16)
                temp16_23(site) = temp16_23(site) + tempC(site)
            Next i
            
            For i = 24 To 31
                tempD(site) = TempArray(i) * 2 ^ (i - 24)
                temp24_31(site) = temp24_31(site) + tempD(site)
            Next i
                                                                                     
            DspArray(0) = temp24_31(site)
            DspArray(1) = temp16_23(site)
            DspArray(2) = temp8_15(site)
            DspArray(3) = temp0_7(site)
                
        End If

        'WaveDefinitions
        DspSrcWave.Data = DspArray
        'DspSrcWave.Plot
        Call TheExec.WaveDefinitions.CreateWaveDefinition("SrcData", DspSrcWave, True)

        'Insert DSSC Loading & Pattern Loading & Pattern Execution.
        With TheHdw.DSSC.Pins("SPI1_MOSI").Pattern(patt).Source
'        With TheHdw.DSSC.Pins("SPI0_MOSI").Pattern(PatName).Source
            .Signals.Add "DSSC_Src"
            .Signals.DefaultSignal = "DSSC_Src"
            .Signals("DSSC_Src").WaveDefinitionName = "SrcData"
            .Signals.Item("DSSC_Src").SampleSize = 260
            .Signals("DSSC_Src").Amplitude = 1
            If glb_TesterType = "Jaguar" Then .Signals("DSSC_Src").LoadSettings
        End With

        TheHdw.Patterns(patt).Test pfNever, 0
        TheHdw.Digital.Patgen.HaltWait
        
        WaveCount = WaveCount + 1
    Wend
        
    
    Close #1
    
    Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function





Public Function SPIROM_Read_RomCode(RomFileName1 As String, RomSize As Long, CapNumOfTimes As String) As Long

Dim TotalDspArray() As Long
Dim DspArray1() As Long
Dim DspArray2() As Long
Dim Count As Long
Dim Rom_Data As Byte

Dim StartAddr As Long
Dim Cap_1st As String
Dim Cap_2nd As String
Dim Cap_3rd As String
Dim Cap_4th As String

Dim SPI_Binary_Signal1 As New DSPWave
Dim SPI_Binary_Signal2 As New DSPWave
Dim SPI_Binary_Erase1 As New DSPWave
Dim SPI_Binary_Erase2 As New DSPWave

    On Error GoTo ErrorHandler
   

    ReDim DspArray1(RomSize / 4 - 1) As Long
    
    Select Case CapNumOfTimes
        Case "Cap_1st"
            StartAddr = 0
        Case "Cap_2nd"
            StartAddr = 4194304
        Case "Cap_3rd"
            StartAddr = 8388608
        Case "Cap_4th"
            StartAddr = 12582912
    End Select
    
    
    Total_Count = StartAddr
    
    
    For Count = 0 To RomSize / 4 - 1
        DspArray1(Count) = 2 ^ 8 - 1
    Next Count
    
    SPI_Binary_Erase1.Data = DspArray1
    SPI_Binary_32MegErase.Data = SPI_Binary_Erase1.Data
    

    Open RomFileName1 For Binary As #1
                      
        ' Create data array for reference
        Count = 0
        While ((Not EOF(1)) And (Count < RomSize / 4))
            Get #1, Total_Count + 1, Rom_Data
            If EOF(1) = False Then
                While ((Not EOF(1)) And (Count < RomSize / 4))
                    Get #1, Total_Count + 1, Rom_Data
                    DspArray1(Count) = Rom_Data
                    Count = Count + 1
                    Total_Count = Total_Count + 1
                    'Debug.Print Rom_Data
                Wend
            End If
        Wend
        SPI_Binary_Signal1.Data = DspArray1
        SPI_Binary_32MegCode.Data = SPI_Binary_Signal1.Data
    
    Close #1
    
'''   Select Case CapNumOfTimes
'''        Case "Cap_1st"
'''            SPI_Binary_32MegCode_PP = SPI_Binary_32MegCode
'''        Case "Cap_2nd"
'''            SPI_Binary_32MegCode_PP = SPI_Binary_32MegCode_PP.Concatenate(SPI_Binary_32MegCode)
'''        Case "Cap_3rd"
'''            SPI_Binary_32MegCode_PP = SPI_Binary_32MegCode_PP.Concatenate(SPI_Binary_32MegCode)
'''        Case "Cap_4th"
'''            SPI_Binary_32MegCode_PP = SPI_Binary_32MegCode_PP.Concatenate(SPI_Binary_32MegCode)
'''    End Select
    
    
    SPIROM_Read_RomCode = Total_Count
    Exit Function
    
    
ErrorHandler:
    If AbortTest Then Exit Function Else Resume Next
    
End Function





Public Function SPIROM_Read_RomCode_32M(RomFileName1 As String, RomSize As Long, CapNumOfTimes As String) As Long

Dim TotalDspArray() As Long
Dim DspArray1() As Long
Dim DspArray2() As Long
Dim Count As Long
Dim Rom_Data As Byte

Dim StartAddr As Long
Dim Cap_1st As String
Dim Cap_2nd As String
Dim Cap_3rd As String
Dim Cap_4th As String
Dim Cap_5th As String
Dim Cap_6th As String
Dim Cap_7th As String
Dim Cap_8th As String

Dim SPI_Binary_Signal1 As New DSPWave
Dim SPI_Binary_Signal2 As New DSPWave
Dim SPI_Binary_Erase1 As New DSPWave
Dim SPI_Binary_Erase2 As New DSPWave

    On Error GoTo ErrorHandler
   

    ReDim DspArray1(RomSize / 8 - 1) As Long
    
    Select Case CapNumOfTimes
        Case "Cap_1st"
            StartAddr = 0
        Case "Cap_2nd"
            StartAddr = 4194304
        Case "Cap_3rd"
            StartAddr = 8388608
        Case "Cap_4th"
            StartAddr = 12582912
        Case "Cap_5th"
            StartAddr = 16777216
        Case "Cap_6th"
            StartAddr = 20971520
        Case "Cap_7th"
            StartAddr = 25165824
        Case "Cap_8th"
            StartAddr = 29360128
    End Select
    
    
    Total_Count = StartAddr
    
    
    For Count = 0 To RomSize / 8 - 1
        DspArray1(Count) = 2 ^ 8 - 1
    Next Count
    
    SPI_Binary_Erase1.Data = DspArray1
    SPI_Binary_32MegErase.Data = SPI_Binary_Erase1.Data
    

    Open RomFileName1 For Binary As #1
                      
        ' Create data array for reference
        Count = 0
        While ((Not EOF(1)) And (Count < RomSize / 8))
            Get #1, Total_Count + 1, Rom_Data
            If EOF(1) = False Then
                While ((Not EOF(1)) And (Count < RomSize / 8))
                    Get #1, Total_Count + 1, Rom_Data
                    DspArray1(Count) = Rom_Data
                    Count = Count + 1
                    Total_Count = Total_Count + 1
                    'Debug.Print Rom_Data
                Wend
            End If
        Wend
        SPI_Binary_Signal1.Data = DspArray1
        SPI_Binary_32MegCode.Data = SPI_Binary_Signal1.Data
    
    Close #1
    
'''   Select Case CapNumOfTimes
'''        Case "Cap_1st"
'''            SPI_Binary_32MegCode_PP = SPI_Binary_32MegCode
'''        Case "Cap_2nd"
'''            SPI_Binary_32MegCode_PP = SPI_Binary_32MegCode_PP.Concatenate(SPI_Binary_32MegCode)
'''        Case "Cap_3rd"
'''            SPI_Binary_32MegCode_PP = SPI_Binary_32MegCode_PP.Concatenate(SPI_Binary_32MegCode)
'''        Case "Cap_4th"
'''            SPI_Binary_32MegCode_PP = SPI_Binary_32MegCode_PP.Concatenate(SPI_Binary_32MegCode)
'''    End Select
    
    
    SPIROM_Read_RomCode_32M = Total_Count
    Exit Function
    
    
ErrorHandler:
    If AbortTest Then Exit Function Else Resume Next
    
End Function




Public Function InitialRead_VBT(PatName As Pattern, DSSCCapPin As PinList, RelayOnPins As PinList, RelayOffPins As PinList) As Long
Dim X As Long
Dim RomFileName1 As String
Dim Var_1 As Integer
Dim SPIROM_code_version As String
Dim SPI_Flag As New SiteLong
Dim site As Variant

    ' 32MByte = 33554421  Hardcoded value for TTR. Set enough buffer
    ' to prevent runtime error array due to actual file size
'    x = 33554434
    X = 16777216



    If TheExec.Flow.EnableWord("SPIROM_1_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).Value
      gS_SPI_Version = RomFileName1

       SPIROM_CheckSum PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If TheExec.Flow.EnableWord("SPIROM_2_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).Value
      gS_SPI_Version = RomFileName1
      
       SPIROM_CheckSum PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
     If TheExec.Flow.EnableWord("SPIROM_3_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).Value
      gS_SPI_Version = RomFileName1
      
       SPIROM_CheckSum PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
   
    If CheckSumResult16M = 1 Then
            For Each site In TheExec.sites
                    SPI_Flag = 1
                    TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
            Next site
            ''control whether the program execute again or not.
            TheExec.Datalog.WriteComment ("print: SPIROM version check pass with " & gS_SPI_Version)
            TheExec.Flow.EnableWord("Write_SPIROM") = False
            TheHdw.Utility.Pins(RelayOffPins).State = tlUtilBitOff
    Else
            For Each site In TheExec.sites
                SPI_Flag = 0
                TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
            Next site
            TheExec.Datalog.WriteComment ("print: SPIROM version check fail with " & gS_SPI_Version)
'''            MsgBox ("The SPIROM function is not successful.")
    End If

End Function


Public Function InitialRead_32M_VBT(PatName As Pattern, DSSCCapPin As PinList, RelayOnPins As PinList, RelayOffPins As PinList) As Long
Dim X As Long
Dim RomFileName1 As String
Dim Var_1 As Integer
Dim SPIROM_code_version As String
Dim SPI_Flag As New SiteLong
Dim site As Variant

    ' 32MByte = 33554421  Hardcoded value for TTR. Set enough buffer
    ' to prevent runtime error array due to actual file size
    X = 33554432
'    x = 16777216

    'Offline simulation
    
    
    
    
    
    
    
    
    '14th, Oct 2019
       'After Discussing with customer,TER-Fred modified for TP checking
    '-------------------------------------------------------------------Strat
    If TheExec.TesterMode = testModeOffline And TheExec.Flow.EnableWord("SPIROMOffline") = True Then
        For Each site In TheExec.sites
            SPI_Flag = 1
            TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
        Next site
        Exit Function
    End If
    '-------------------------------------------------------------------End
    
    
    
    
    
    
    
    
    If TheExec.Flow.EnableWord("SPIROM_1_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).Value
      gS_SPI_Version = RomFileName1

       SPIROM_CheckSum_32M PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If TheExec.Flow.EnableWord("SPIROM_2_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).Value
      gS_SPI_Version = RomFileName1
      
       SPIROM_CheckSum_32M PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If TheExec.Flow.EnableWord("SPIROM_3_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).Value
      gS_SPI_Version = RomFileName1
      
       SPIROM_CheckSum_32M PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If CheckSumResult32M = 1 Then
            For Each site In TheExec.sites
                SPI_Flag = 1
                TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
            Next site
             ''control whether the program execute again or not.
            TheExec.Datalog.WriteComment ("print: SPIROM version check pass with " & gS_SPI_Version)
            TheExec.Flow.EnableWord("Write_SPIROM") = False
             TheHdw.Utility.Pins(RelayOffPins).State = tlUtilBitOff
    Else
            For Each site In TheExec.sites
                SPI_Flag = 0
                TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
            Next site
            TheExec.Datalog.WriteComment ("print: SPIROM version check fail with " & gS_SPI_Version)
'''            MsgBox ("The SPIROM function is not successful.")
    End If
  

End Function



Public Function InitialRead_32M_VBT_SectorErase(PatName As Pattern, DSSCCapPin As PinList, RelayOnPins As PinList, RelayOffPins As PinList) As Long
Dim X As Long
Dim RomFileName1 As String
Dim Var_1 As Integer
Dim SPIROM_code_version As String
Dim SPI_Flag As New SiteLong
Dim site As Variant

    ' 32MByte = 33554421  Hardcoded value for TTR. Set enough buffer
    ' to prevent runtime error array due to actual file size
    X = 33554432
'    x = 16777216

    'Offline simulation
    
    

       '14th, Oct 2019
       'After Discussing with customer,TER-Fred modified for TP checking
        '-------------------------------------------------------------------Strat
        If TheExec.TesterMode = testModeOffline And TheExec.Flow.EnableWord("SPIROMOffline") = True Then
            For Each site In TheExec.sites
                SPI_Flag = 1
                TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
            Next site
            Exit Function
        End If
        '-------------------------------------------------------------------End
        
    
    
    
    
    
    If TheExec.Flow.EnableWord("SPIROM_1_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).Value
      gS_SPI_Version = RomFileName1

       SPIROM_CheckSum_32M_SectorErase PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If TheExec.Flow.EnableWord("SPIROM_2_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).Value
      gS_SPI_Version = RomFileName1
      
       SPIROM_CheckSum_32M_SectorErase PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If TheExec.Flow.EnableWord("SPIROM_3_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).Value
      gS_SPI_Version = RomFileName1
      
       SPIROM_CheckSum_32M_SectorErase PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If CheckSumResult32M = 1 Then
            For Each site In TheExec.sites
                SPI_Flag = 1
                TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
            Next site
             ''control whether the program execute again or not.
            TheExec.Datalog.WriteComment ("print: SPIROM version check pass with " & gS_SPI_Version)
            TheExec.Flow.EnableWord("Write_SPIROM_SectorErase") = False
             TheHdw.Utility.Pins(RelayOffPins).State = tlUtilBitOff
    Else
            For Each site In TheExec.sites
                SPI_Flag = 0
                TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_First"
            Next site
            TheExec.Datalog.WriteComment ("print: SPIROM version check fail with " & gS_SPI_Version)
'''            MsgBox ("The SPIROM function is not successful.")
    End If
  

End Function











Public Function CheckSum_VBT(PatName As Pattern, DSSCCapPin As PinList, RelayOnPins As PinList, RelayOffPins As PinList) As Long
Dim X As Long
Dim RomFileName1 As String
Dim Var_1 As Integer
Dim SPIROM_code_version As String
Dim SPI_Flag As New SiteLong
Dim site As Variant


    ' 32MByte = 33554421  Hardcoded value for TTR. Set enough buffer
    ' to prevent runtime error array due to actual file size
'    x = 33554434
    X = 16777216

    If TheExec.Flow.EnableWord("SPIROM_1_Write") = True Then
            RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).Value
            gS_SPI_Version = RomFileName1
            SPIROM_CheckSum PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
        If TheExec.Flow.EnableWord("SPIROM_2_Write") = True Then
            RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).Value
            gS_SPI_Version = RomFileName1
            SPIROM_CheckSum PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If TheExec.Flow.EnableWord("SPIROM_3_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).Value
      gS_SPI_Version = RomFileName1
      
       SPIROM_CheckSum_32M PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
        If CheckSumResult16M = 1 Then
            For Each site In TheExec.sites
                        SPI_Flag = 1
                        TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_Second"
            Next site
            ''control whether the program execute again or not.
            TheExec.Datalog.WriteComment ("print: Auto-trim finished, SPIROM version check pass " & gS_SPI_Version)
            TheExec.Flow.EnableWord("Write_SPIROM") = False
            TheHdw.Utility.Pins(RelayOffPins).State = tlUtilBitOff
        Else
             For Each site In TheExec.sites
                       SPI_Flag = 0
                       TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_Second"
            Next site
            'SPI_Version = SPIROM_code_version & " checksum failed"
            TheExec.Datalog.WriteComment ("print: Auto-trim finished, SPIROM version check fail " & gS_SPI_Version)
            MsgBox ("The SPIROM function is not successful.")
        End If

End Function




Public Function CheckSum_32M_VBT(PatName As Pattern, DSSCCapPin As PinList, RelayOnPins As PinList, RelayOffPins As PinList) As Long
Dim X As Long
Dim RomFileName1 As String
Dim Var_1 As Integer
Dim SPIROM_code_version As String
Dim SPI_Flag As New SiteLong
Dim site As Variant


    ' 32MByte = 33554421  Hardcoded value for TTR. Set enough buffer
    ' to prevent runtime error array due to actual file size
    X = 33554432
'    x = 16777216

    If TheExec.Flow.EnableWord("SPIROM_1_Write") = True Then
            RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).Value
            gS_SPI_Version = RomFileName1
            SPIROM_CheckSum_32M PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If TheExec.Flow.EnableWord("SPIROM_2_Write") = True Then
            RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).Value
            gS_SPI_Version = RomFileName1
            SPIROM_CheckSum_32M PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If TheExec.Flow.EnableWord("SPIROM_3_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).Value
      gS_SPI_Version = RomFileName1
      
       SPIROM_CheckSum_32M PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If CheckSumResult32M = 1 Then
            For Each site In TheExec.sites
                    SPI_Flag = 1
                    TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_Second"
            Next site
            ''control whether the program execute again or not.
            TheExec.Datalog.WriteComment ("print: Auto-trim finished, SPIROM version check pass " & gS_SPI_Version)
            TheExec.Flow.EnableWord("Write_SPIROM") = False
            TheHdw.Utility.Pins(RelayOffPins).State = tlUtilBitOff
    Else
            For Each site In TheExec.sites
                    SPI_Flag = 0
                    TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_Second"
            Next site
            TheExec.Datalog.WriteComment ("print: Auto-trim finished, SPIROM version check fail " & gS_SPI_Version)
            'SPI_Version = SPIROM_code_version & " checksum failed"
            
            
            
            
            
            '21st, Oct 2019
            'After Discussing with customer,TER-Fred modified for TP checking.
            '=======================================Start
'            MsgBox ("The SPIROM function is not successful.")                    'Original code without modification
            TheExec.AddOutput "The SPIROM function is not successful."
               '=======================================End
    End If
    

End Function


Public Function CheckSum_32M_VBT_SectorErase(PatName As Pattern, DSSCCapPin As PinList, RelayOnPins As PinList, RelayOffPins As PinList) As Long
Dim X As Long
Dim RomFileName1 As String
Dim Var_1 As Integer
Dim SPIROM_code_version As String
Dim SPI_Flag As New SiteLong
Dim site As Variant

    X = 33554432
    
    If TheExec.Flow.EnableWord("SPIROM_1_Write") = True Then
            RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).Value
            gS_SPI_Version = RomFileName1
            SPIROM_CheckSum_32M_SectorErase PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If TheExec.Flow.EnableWord("SPIROM_2_Write") = True Then
            RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).Value
            gS_SPI_Version = RomFileName1
            SPIROM_CheckSum_32M_SectorErase PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If TheExec.Flow.EnableWord("SPIROM_3_Write") = True Then
            RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).Value
            gS_SPI_Version = RomFileName1
            SPIROM_CheckSum_32M_SectorErase PatName, DSSCCapPin, RomFileName1, X, RelayOnPins
    End If
    
    If CheckSumResult32M = 1 Then
            For Each site In TheExec.sites
                    SPI_Flag = 1
                    TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_Second"
            Next site
            ''control whether the program execute again or not.
            TheExec.Datalog.WriteComment ("print: Auto-trim finished, SPIROM version check pass " & gS_SPI_Version)
            TheExec.Flow.EnableWord("Write_SPIROM_SectorErase") = False
            TheHdw.Utility.Pins(RelayOffPins).State = tlUtilBitOff
    Else
            For Each site In TheExec.sites
                    SPI_Flag = 0
                    TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "CheckSum_Second"
            Next site
            TheExec.Datalog.WriteComment ("print: Auto-trim finished, SPIROM version check fail " & gS_SPI_Version)
            'SPI_Version = SPIROM_code_version & " checksum failed"

            
            
            
            '21st, Oct 2019
            'After Discussing with customer,TER-Fred modified for TP checking.
            '=======================================Start
'            MsgBox ("The SPIROM function is not successful.")                    'Original code without modification
            TheExec.AddOutput "The SPIROM function is not successful."
               '=======================================End
    End If
    

End Function




Public Function SPIROM_Program_VBT(PatName As Pattern, RelayOnPins As PinList) As Long

Dim X As Long
Dim RomFileName1 As String

    If TheExec.Flow.EnableWord("SPIROM_1_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).Value
        
        TheExec.Datalog.WriteComment ("")
        TheExec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call SPIROM_FL_PP_VBT(PatName, RomFileName1, RelayOnPins)
        TheExec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    End If
    
     If TheExec.Flow.EnableWord("SPIROM_2_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).Value
        
        TheExec.Datalog.WriteComment ("")
        TheExec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call SPIROM_FL_PP_VBT(PatName, RomFileName1, RelayOnPins)
        TheExec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    End If
    
     If TheExec.Flow.EnableWord("SPIROM_3_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).Value
        
        TheExec.Datalog.WriteComment ("")
        TheExec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call SPIROM_FL_PP_VBT(PatName, RomFileName1, RelayOnPins)
        TheExec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    End If
    
'    If TheExec.Sites.site(0).FlagState("ate2-20130905-V1.2.920.spi") = 1 Then
 '       RomFileName1 = ".\ate2-20130905-V1.2.920.spi"
  '      Call SPIROM_FL_PP_VBT(PatName, RomFileName1)
   '     TheExec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
   ' End If
    
    
'    If TheExec.Sites.site(0).FlagState("ate2_20130924_V1.2.974.spi") = 1 Then
 '       RomFileName1 = ".\ate2_20130924_V1.2.974.spi"
  '      Call SPIROM_FL_PP_VBT(PatName, RomFileName1)
   '     TheExec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
   ' End If

End Function




Public Function SPIROM_Program_32M_VBT(PatName As Pattern, RelayOnPins As PinList) As Long

Dim X As Long
Dim RomFileName1 As String

    If TheExec.Flow.EnableWord("SPIROM_1_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(1, 2).Value
        
        TheExec.Datalog.WriteComment ("")
        TheExec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call SPIROM_FL_PP_VBT_32MByte(PatName, RomFileName1, RelayOnPins)
        TheExec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    End If
    
     If TheExec.Flow.EnableWord("SPIROM_2_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(2, 2).Value
        
        TheExec.Datalog.WriteComment ("")
        TheExec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call SPIROM_FL_PP_VBT_32MByte(PatName, RomFileName1, RelayOnPins)
        TheExec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    End If

    If TheExec.Flow.EnableWord("SPIROM_3_Write") = True Then
      RomFileName1 = Worksheets("SpiromCodeFile").Cells(3, 2).Value
        
        TheExec.Datalog.WriteComment ("")
        TheExec.Datalog.WriteComment (RomFileName1 + " start to program.")
        Call SPIROM_FL_PP_VBT_32MByte(PatName, RomFileName1, RelayOnPins)
        TheExec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
    End If
   
'    If TheExec.Sites.site(0).FlagState("ate2-20130905-V1.2.920.spi") = 1 Then
 '       RomFileName1 = ".\ate2-20130905-V1.2.920.spi"
  '      Call SPIROM_FL_PP_VBT(PatName, RomFileName1)
   '     TheExec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
   ' End If
    
    
'    If TheExec.Sites.site(0).FlagState("ate2_20130924_V1.2.974.spi") = 1 Then
 '       RomFileName1 = ".\ate2_20130924_V1.2.974.spi"
  '      Call SPIROM_FL_PP_VBT(PatName, RomFileName1)
   '     TheExec.Datalog.WriteComment (RomFileName1 + " has been programmed.")
   ' End If

End Function






Public Function SPIROM_InitRead(PatName As Pattern, DSSCCapPin As PinList, RomSize As Long, RomFileName1 As String, StartingAddr As Integer, RelayOnPins As PinList) As Long
    
Dim Rom_Data As Byte
Dim TotalDspArray() As Long
Dim DspArray1() As Long
Dim DspArray2() As Long
Dim Count As Long
Dim SrcWaveName As String
Dim StartSTR As String
Dim result As New SiteDouble
Dim WaveCount As Long
Dim SrcDataSize As Long
Dim DSSCCapSize As Long
Dim DspWaveCap As New DSPWave
Dim DSSCCap_MByte_Signal As New DSPWave
Dim DSSCCap_All_Signal_SUM As New DSPWave

Dim SPI_Binary_Signal1 As New DSPWave
Dim SPI_Binary_Signal2 As New DSPWave


    On Error GoTo ErrorHandler
    PatName = Worksheets("Pattern_Sets_SPIROM").Cells(7, 5).Value
    
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn

    ' Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.5

    If glb_TesterType = "Jaguar" Then TheHdw.Digital.Patterns.Pat(PatName).Load
    TheHdw.Digital.Patgen.TimeOut = 800
   
   
    DSSCCapSize = 4194304
    DspWaveCap.Clear
    SPI_Binary_Signal.Clear
                       
            
    Dim AddrIndx As Integer
    Dim AddrSize As Integer
    
    ' If memory size is 32MByte, pattern read address inc is 1MByte
    ' therefore the total iteration: AddrSize= 32
'    AddrSize = 8
''    AddrSize = 4
       
    AddrIndx = StartingAddr
'''    Do While AddrIndx <> AddrSize
    
        DSSCCap_MByte_Signal.Clear
        With TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture
            .Signals.Add "DSSC_Cap_All"
            If glb_TesterType = "Jaguar" Then .Signals("DSSC_Cap_All").Offset = 0
            .Signals.Item("DSSC_Cap_All").SampleSize = DSSCCapSize
            .Signals("DSSC_Cap_All").LoadSettings
        End With
        TheHdw.Wait 0.05
    
        StartSTR = "M32Mbyte_4MegCap_Addr" & CStr(AddrIndx)
        TheHdw.Patterns(PatName).start (StartSTR)
        TheHdw.Digital.Patgen.HaltWait
             
        DSSCCap_MByte_Signal = TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture.Signals("DSSC_Cap_All").DSPWave
        
        ' For debug verify captured segment
        'If AddrIndx = 0 Or AddrIndx = 7 Then
        '    DSSCCap_MByte_Signal.Plot "Read Data Result " & AddrIndx
        '    SPI_Bin_MByte = SPI_Binary_Signal.Select(DSSCCapSize * AddrIndx, 1, DSSCCapSize)
        'End If

'''        If AddrIndx = 0 Then
            DspWaveCap = DSSCCap_MByte_Signal
'''        Else
'''            DspWaveCap = DspWaveCap.Concatenate(DSSCCap_MByte_Signal)
'''        End If
           
''        AddrIndx = AddrIndx + 1
        
'''    Loop
    
    
    '==========================================================
    
    TheHdw.Wait 0.05
    
    Dim Different_Value As New DSPWave
    Dim Abs_Different_Value As New DSPWave
    Dim IndexOfMaximumValue As Long
    Dim IndexOfMinimumValue As Long
    Dim MaximumValue As Double
    Dim MinimumValue As Double
'    Dim SPI_Flag As New SiteLong
    Dim site As Variant
    Dim Sum As New DSPWave
    
'    SPI_Flag_0 = False
'    SPI_Flag_1 = False
'    SPI_Flag_2 = False
'    SPI_Flag_3 = False
    
    For Each site In TheExec.sites
        Sum = DspWaveCap
        Different_Value = SPI_Binary_32MegCode.Subtract(Sum)
        Abs_Different_Value = Different_Value.Abs
        MaximumValue = Abs_Different_Value.CalcMaximumValue(IndexOfMaximumValue)
        If MaximumValue = 0 Then
            Select Case StartingAddr
                Case 0
                        SPI_Flag_0 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                        SPI_Flag_temp0 = 1
                Case 1
                        SPI_Flag_1 = True     'SPI-ROM PROGRAMMED with current ROM CODE
                        SPI_Flag_temp1 = 1
                Case 2
                        SPI_Flag_2 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                        SPI_Flag_temp2 = 1
               Case 3
                        SPI_Flag_3 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                        SPI_Flag_temp3 = 1
          End Select
        Else
            Select Case StartingAddr
                Case 0
                        SPI_Flag_0 = False     'SPI-ROM Erase
                        SPI_Flag_temp0 = 2
               Case 1
                        SPI_Flag_1 = False     'SPI-ROM Erase
                        SPI_Flag_temp1 = 2
                Case 2
                        SPI_Flag_2 = False      'SPI-ROM Erase
                        SPI_Flag_temp2 = 2
                Case 3
                        SPI_Flag_3 = False      'SPI-ROM Erase
                        SPI_Flag_temp3 = 2
           End Select
        
        End If
    Next site
        
        

        
    '============================================================
        
    Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function




Public Function SPIROM_CheckSum(PatName As Pattern, DSSCCapPin As PinList, RomFileName1 As String, X As Long, RelayOnPins As PinList) As Long
    
Dim Cap_1st As String
Dim Cap_2nd As String
Dim Cap_3rd As String
Dim Cap_4th As String


    On Error GoTo ErrorHandler

    SPIROM_Read_RomCode RomFileName1, X, "Cap_1st"
    Call SPIROM_InitRead(PatName, DSSCCapPin, X, RomFileName1, 0, RelayOnPins)
    SPIROM_Read_RomCode RomFileName1, X, "Cap_2nd"
    Call SPIROM_InitRead(PatName, DSSCCapPin, X, RomFileName1, 1, RelayOnPins)
    SPIROM_Read_RomCode RomFileName1, X, "Cap_3rd"
    Call SPIROM_InitRead(PatName, DSSCCapPin, X, RomFileName1, 2, RelayOnPins)
    SPIROM_Read_RomCode RomFileName1, X, "Cap_4th"
    Call SPIROM_InitRead(PatName, DSSCCapPin, X, RomFileName1, 3, RelayOnPins)

    '============================================================
    For Each site In TheExec.sites
            SPI_Flag_Sum_16M = SPI_Flag_temp0 And SPI_Flag_temp1 And SPI_Flag_temp2 And SPI_Flag_temp3
    Next site

        CheckSumResult16M = 15
        For Each site In TheExec.sites
            CheckSumResult16M = SPI_Flag_Sum_16M(site) And 15 And CheckSumResult16M
        Next site
         
         
    Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function SPIROM_InitRead_32M(PatName As Pattern, DSSCCapPin As PinList, RomSize As Long, RomFileName1 As String, StartingAddr As Integer, RelayOnPins As PinList) As Long
    
Dim Rom_Data As Byte
Dim TotalDspArray() As Long
Dim DspArray1() As Long
Dim DspArray2() As Long
Dim Count As Long
Dim SrcWaveName As String
Dim StartSTR As String
Dim result As New SiteDouble
Dim WaveCount As Long
Dim SrcDataSize As Long
Dim DSSCCapSize As Long
Dim DspWaveCap As New DSPWave
Dim DSSCCap_MByte_Signal As New DSPWave
Dim DSSCCap_All_Signal_SUM As New DSPWave

Dim SPI_Binary_Signal1 As New DSPWave
Dim SPI_Binary_Signal2 As New DSPWave
Dim PatCount As Long
Dim PattArray() As String
Dim patt As String

    'On Error GoTo ErrorHandler !!!!
    
'''    PatName = Worksheets("PatSets_SPIROM").Cells(10, 5).Value
     
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn

    ' Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.5

    Call PATT_GetPatListFromPatternSet(PatName.Value, PattArray, PatCount)
    patt = CStr(PattArray(0))
    If glb_TesterType = "Jaguar" Then TheHdw.Digital.Patterns.Pat(patt).Load
    TheHdw.Digital.Patgen.TimeOut = 10
    
    DSSCCapSize = 4194304
    DspWaveCap.Clear
    SPI_Binary_Signal.Clear
                       
            
    Dim AddrIndx As Integer
    Dim AddrSize As Integer
    
    ' If memory size is 32MByte, pattern read address inc is 1MByte
    ' therefore the total iteration: AddrSize= 32
'    AddrSize = 8
''    AddrSize = 4
       
    AddrIndx = StartingAddr
'''    Do While AddrIndx <> AddrSize
    
        DSSCCap_MByte_Signal.Clear
        With TheHdw.DSSC.Pins(DSSCCapPin).Pattern(patt).Capture
'        With TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture
            .Signals.Add "DSSC_Cap_All"
            If glb_TesterType = "Jaguar" Then .Signals("DSSC_Cap_All").Offset = 0
            .Signals.Item("DSSC_Cap_All").SampleSize = DSSCCapSize
            .Signals("DSSC_Cap_All").LoadSettings
        End With
        TheHdw.Wait 0.05
    
        StartSTR = "M32Mbyte_4MegCap_32MAddr" & CStr(AddrIndx)
        TheHdw.Patterns(patt).start (StartSTR)
        TheHdw.Digital.Patgen.HaltWait
             
        DSSCCap_MByte_Signal = TheHdw.DSSC.Pins(DSSCCapPin).Pattern(patt).Capture.Signals("DSSC_Cap_All").DSPWave
        
        ' For debug verify captured segment
        'If AddrIndx = 0 Or AddrIndx = 7 Then
        '    DSSCCap_MByte_Signal.Plot "Read Data Result " & AddrIndx
        '    SPI_Bin_MByte = SPI_Binary_Signal.Select(DSSCCapSize * AddrIndx, 1, DSSCCapSize)
        'End If

'''        If AddrIndx = 0 Then
            DspWaveCap = DSSCCap_MByte_Signal
'''        Else
'''            DspWaveCap = DspWaveCap.Concatenate(DSSCCap_MByte_Signal)
'''        End If
           
''        AddrIndx = AddrIndx + 1
        
'''    Loop
    
    
    '==========================================================
    
    TheHdw.Wait 0.05
    
    Dim Different_Value As New DSPWave
    Dim Abs_Different_Value As New DSPWave
    Dim IndexOfMaximumValue As Long
    Dim IndexOfMinimumValue As Long
    Dim MaximumValue As Double
    Dim MinimumValue As Double
'    Dim SPI_Flag As New SiteLong
    Dim site As Variant
    Dim Sum As New DSPWave
    
'    SPI_Flag_0 = False
'    SPI_Flag_1 = False
'    SPI_Flag_2 = False
'    SPI_Flag_3 = False
    
    For Each site In TheExec.sites
        Sum = DspWaveCap
        Different_Value = SPI_Binary_32MegCode.Subtract(Sum)
        Abs_Different_Value = Different_Value.Abs
        MaximumValue = Abs_Different_Value.CalcMaximumValue(IndexOfMaximumValue)
        If MaximumValue = 0 Then
            Select Case StartingAddr
            Case 0
                    SPI_Flag_0 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                    SPI_Flag_temp0 = 1
            Case 1
                    SPI_Flag_1 = True     'SPI-ROM PROGRAMMED with current ROM CODE
                    SPI_Flag_temp1 = 1
            Case 2
                    SPI_Flag_2 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                    SPI_Flag_temp2 = 1
            Case 3
                    SPI_Flag_3 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                    SPI_Flag_temp3 = 1
            Case 4
                    SPI_Flag_4 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                    SPI_Flag_temp4 = 1
            Case 5
                    SPI_Flag_5 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                    SPI_Flag_temp5 = 1
            Case 6
                    SPI_Flag_6 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                    SPI_Flag_temp6 = 1
            Case 7
                    SPI_Flag_7 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                    SPI_Flag_temp7 = 1
           End Select
        Else
            Select Case StartingAddr
            Case 0
                    SPI_Flag_0 = False     'SPI-ROM Erase
                    SPI_Flag_temp0 = 2
            Case 1
                    SPI_Flag_1 = False     'SPI-ROM Erase
                    SPI_Flag_temp1 = 2
            Case 2
                    SPI_Flag_2 = False      'SPI-ROM Erase
                    SPI_Flag_temp2 = 2
            Case 3
                    SPI_Flag_3 = False      'SPI-ROM Erase
                    SPI_Flag_temp3 = 2
            Case 4
                    SPI_Flag_4 = False      'SPI-ROM Erase
                    SPI_Flag_temp4 = 2
            Case 5
                    SPI_Flag_5 = False      'SPI-ROM Erase
                    SPI_Flag_temp5 = 2
            Case 6
                    SPI_Flag_6 = False      'SPI-ROM Erase
                    SPI_Flag_temp6 = 2
            Case 7
                    SPI_Flag_7 = False      'SPI-ROM Erase
                    SPI_Flag_temp7 = 2
           End Select
        
        End If
    Next site
        
    '============================================================
        
    Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function




Public Function SPIROM_InitRead_32M_SectorErase(PatName As Pattern, DSSCCapPin As PinList, RomSize As Long, RomFileName1 As String, StartingAddr As Integer, RelayOnPins As PinList) As Long
    
Dim Rom_Data As Byte
Dim TotalDspArray() As Long
Dim DspArray1() As Long
Dim DspArray2() As Long
Dim Count As Long
Dim SrcWaveName As String
Dim StartSTR As String
Dim result As New SiteDouble
Dim WaveCount As Long
Dim SrcDataSize As Long
Dim DSSCCapSize As Long
Dim DspWaveCap As New DSPWave
Dim DSSCCap_MByte_Signal As New DSPWave
Dim DSSCCap_All_Signal_SUM As New DSPWave

Dim SPI_Binary_Signal1 As New DSPWave
Dim SPI_Binary_Signal2 As New DSPWave
Dim PatCount As Long
Dim PattArray() As String
Dim patt As String

    'On Error GoTo ErrorHandler !!!!
    
'''    PatName = Worksheets("PatSets_SPIROM").Cells(10, 5).Value
     
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn

    ' Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.5

    Call PATT_GetPatListFromPatternSet(PatName.Value, PattArray, PatCount)
    patt = CStr(PattArray(0))
    If glb_TesterType = "Jaguar" Then TheHdw.Digital.Patterns.Pat(patt).Load
    TheHdw.Digital.Patgen.TimeOut = 10
    
    DSSCCapSize = 4194304
    DspWaveCap.Clear
    SPI_Binary_Signal.Clear
                       
            
    Dim AddrIndx As Integer
    Dim AddrSize As Integer
    
    ' If memory size is 32MByte, pattern read address inc is 1MByte
    ' therefore the total iteration: AddrSize= 32
'    AddrSize = 8
''    AddrSize = 4
       
    AddrIndx = StartingAddr
'''    Do While AddrIndx <> AddrSize
    
        DSSCCap_MByte_Signal.Clear
        With TheHdw.DSSC.Pins(DSSCCapPin).Pattern(patt).Capture
'        With TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture
            .Signals.Add "DSSC_Cap_All"
            If glb_TesterType = "Jaguar" Then .Signals("DSSC_Cap_All").Offset = 0
            .Signals.Item("DSSC_Cap_All").SampleSize = DSSCCapSize
            .Signals("DSSC_Cap_All").LoadSettings
        End With
        TheHdw.Wait 0.05
    
        StartSTR = "M32Mbyte_4MegCap_32MAddr" & CStr(AddrIndx)
        TheHdw.Patterns(patt).start (StartSTR)
        TheHdw.Digital.Patgen.HaltWait
             
        DSSCCap_MByte_Signal = TheHdw.DSSC.Pins(DSSCCapPin).Pattern(patt).Capture.Signals("DSSC_Cap_All").DSPWave
        
        ' For debug verify captured segment
        'If AddrIndx = 0 Or AddrIndx = 7 Then
        '    DSSCCap_MByte_Signal.Plot "Read Data Result " & AddrIndx
        '    SPI_Bin_MByte = SPI_Binary_Signal.Select(DSSCCapSize * AddrIndx, 1, DSSCCapSize)
        'End If

'''        If AddrIndx = 0 Then
            DspWaveCap = DSSCCap_MByte_Signal
'''        Else
'''            DspWaveCap = DspWaveCap.Concatenate(DSSCCap_MByte_Signal)
'''        End If
           
''        AddrIndx = AddrIndx + 1
        
'''    Loop
    
    
    '==========================================================
    
    TheHdw.Wait 0.05
    
    Dim Different_Value As New DSPWave
    Dim Abs_Different_Value As New DSPWave
    Dim IndexOfMaximumValue As Long
    Dim IndexOfMinimumValue As Long
    Dim MaximumValue As Double
    Dim MinimumValue As Double
'    Dim SPI_Flag As New SiteLong
    Dim site As Variant
    Dim Sum As New DSPWave
    
'    SPI_Flag_0 = False
'    SPI_Flag_1 = False
'    SPI_Flag_2 = False
'    SPI_Flag_3 = False
    
    For Each site In TheExec.sites
        Sum = DspWaveCap
        Different_Value = SPI_Binary_32MegCode.Subtract(Sum)
        Abs_Different_Value = Different_Value.Abs
        MaximumValue = Abs_Different_Value.CalcMaximumValue(IndexOfMaximumValue)
        If MaximumValue = 0 Then
            Select Case StartingAddr
            Case 0
                    SPI_Flag_0 = True      'SPI-ROM PROGRAMMED with current ROM CODE
                    SPI_Flag_temp0 = 1
           End Select
        Else
            Select Case StartingAddr
            Case 0
                    SPI_Flag_0 = False     'SPI-ROM Erase
                    SPI_Flag_temp0 = 2
           End Select
        
        End If
    Next site
        
    '============================================================
        
    Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function






Public Function SPIROM_CheckSum_32M(PatName As Pattern, DSSCCapPin As PinList, RomFileName1 As String, X As Long, RelayOnPins As PinList) As Long
    
Dim Cap_1st As String
Dim Cap_2nd As String
Dim Cap_3rd As String
Dim Cap_4th As String
Dim Cap_5th As String
Dim Cap_6th As String
Dim Cap_7th As String
Dim Cap_8th As String


    On Error GoTo ErrorHandler

    SPIROM_Read_RomCode_32M RomFileName1, X, "Cap_1st"
    Call SPIROM_InitRead_32M(PatName, DSSCCapPin, X, RomFileName1, 0, RelayOnPins)
    SPIROM_Read_RomCode_32M RomFileName1, X, "Cap_2nd"
    Call SPIROM_InitRead_32M(PatName, DSSCCapPin, X, RomFileName1, 1, RelayOnPins)
    SPIROM_Read_RomCode_32M RomFileName1, X, "Cap_3rd"
    Call SPIROM_InitRead_32M(PatName, DSSCCapPin, X, RomFileName1, 2, RelayOnPins)
    SPIROM_Read_RomCode_32M RomFileName1, X, "Cap_4th"
    Call SPIROM_InitRead_32M(PatName, DSSCCapPin, X, RomFileName1, 3, RelayOnPins)

    SPIROM_Read_RomCode_32M RomFileName1, X, "Cap_5th"
    Call SPIROM_InitRead_32M(PatName, DSSCCapPin, X, RomFileName1, 4, RelayOnPins)
    SPIROM_Read_RomCode_32M RomFileName1, X, "Cap_6th"
    Call SPIROM_InitRead_32M(PatName, DSSCCapPin, X, RomFileName1, 5, RelayOnPins)
    SPIROM_Read_RomCode_32M RomFileName1, X, "Cap_7th"
    Call SPIROM_InitRead_32M(PatName, DSSCCapPin, X, RomFileName1, 6, RelayOnPins)
    SPIROM_Read_RomCode_32M RomFileName1, X, "Cap_8th"
    Call SPIROM_InitRead_32M(PatName, DSSCCapPin, X, RomFileName1, 7, RelayOnPins)

    '============================================================
    For Each site In TheExec.sites
            SPI_Flag_Sum_32M = SPI_Flag_temp0 And SPI_Flag_temp1 And SPI_Flag_temp2 And SPI_Flag_temp3 _
                                                    And SPI_Flag_temp4 And SPI_Flag_temp5 And SPI_Flag_temp6 And SPI_Flag_temp7
    Next site

        CheckSumResult32M = 15
       For Each site In TheExec.sites
        CheckSumResult32M = SPI_Flag_Sum_32M(site) And 15 And CheckSumResult32M
        '     CheckSumResult32M = SPI_Flag_Sum_32M And 15 And CheckSumResult32M
        Next site
    
    Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function




Public Function SPIROM_CheckSum_32M_SectorErase(PatName As Pattern, DSSCCapPin As PinList, RomFileName1 As String, X As Long, RelayOnPins As PinList) As Long
    
Dim Cap_1st As String
Dim Cap_2nd As String
Dim Cap_3rd As String
Dim Cap_4th As String
Dim Cap_5th As String
Dim Cap_6th As String
Dim Cap_7th As String
Dim Cap_8th As String


    On Error GoTo ErrorHandler

    SPIROM_Read_RomCode_32M RomFileName1, X, "Cap_1st"
    Call SPIROM_InitRead_32M_SectorErase(PatName, DSSCCapPin, X, RomFileName1, 0, RelayOnPins)


    '============================================================
    For Each site In TheExec.sites
            SPI_Flag_Sum_32M = SPI_Flag_temp0
    Next site

        CheckSumResult32M = 15
       For Each site In TheExec.sites
            CheckSumResult32M = SPI_Flag_Sum_32M(site) And 15 And CheckSumResult32M
        '     CheckSumResult32M = SPI_Flag_Sum_32M And 15 And CheckSumResult32M
        Next site
    
    Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function










Public Function SPIROM_InitRead_32MByte(PatName As Pattern, DSSCCapPin As PinList, RomSize As Long, RomFileName1 As String) As Long
    
Dim Rom_Data As Byte
Dim TotalDspArray() As Long
Dim DspArray1() As Long
Dim DspArray2() As Long
Dim Count As Long
Dim SrcWaveName As String
Dim StartSTR As String
Dim result As New SiteDouble
Dim WaveCount As Long
Dim SrcDataSize As Long
Dim DSSCCapSize As Long
Dim DspWaveCap As New DSPWave
Dim DSSCCap_MByte_Signal As New DSPWave
Dim DSSCCap_All_Signal_SUM As New DSPWave

Dim SPI_Binary_Signal1 As New DSPWave
Dim SPI_Binary_Signal2 As New DSPWave


    On Error GoTo ErrorHandler

    TheHdw.Utility.Pins("K3_SPI0_SCLK_SSIN_ROM,K5_SPI0_MISO_MOSI_ROM").State = tlUtilBitOn

    ' Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.5

    If glb_TesterType = "Jaguar" Then TheHdw.Digital.Patterns.Pat(PatName).Load
    TheHdw.Digital.Patgen.TimeOut = 800
   
   
    DSSCCapSize = 4194304
    SrcDataSize = RomSize
    Total_Count = 0
    WaveCount = 0

    ReDim DspArray1(RomSize / 2) As Long
    ReDim DspArray2(RomSize / 2) As Long
    
    
    DspWaveCap.Clear
    SPI_Binary_Signal.Clear

    'ReDim Preserve DspArray(RomSize) As Double
    For Count = 0 To RomSize - 1
        If Count <= (RomSize / 2) Then
            DspArray1(Count) = 2 ^ 8 - 1
        Else
            DspArray2(Count - ((RomSize / 2) + 1)) = 2 ^ 8 - 1
        End If
    Next Count
    
    '==========================================================
    
    Open RomFileName1 For Binary As #1
             
             
        ' Create data array for reference
        Count = 0
        While ((Not EOF(1)) And (Count < SrcDataSize))
            Get #1, Total_Count + 1, Rom_Data
            If Count <= (RomSize / 2) Then
                DspArray1(Count) = Rom_Data
            Else
                DspArray2(Count - ((RomSize / 2) + 1)) = Rom_Data
            End If
            Count = Count + 1
            Total_Count = Total_Count + 1
            'Debug.Print Rom_Data
        Wend
        SPI_Binary_Signal1.Data = DspArray1
        SPI_Binary_Signal2.Data = DspArray2
        SPI_Binary_Signal = SPI_Binary_Signal1.Concatenate(SPI_Binary_Signal2)
                
                
        Dim AddrIndx As Integer
        Dim AddrSize As Integer
        
        ' If memory size is 32MByte, pattern read address inc is 1MByte
        ' therefore the total iteration: AddrSize= 32
        AddrSize = 8
           
        AddrIndx = 0
        Do While AddrIndx <> AddrSize
        
            DSSCCap_MByte_Signal.Clear
            With TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture
                .Signals.Add "DSSC_Cap_All"
                If glb_TesterType = "Jaguar" Then .Signals("DSSC_Cap_All").Offset = 0
                .Signals.Item("DSSC_Cap_All").SampleSize = DSSCCapSize
                .Signals("DSSC_Cap_All").LoadSettings
            End With
            TheHdw.Wait 0.05
        
            StartSTR = "M32Mbyte_4MegCap_Addr" & CStr(AddrIndx)
            TheHdw.Patterns(PatName).start (StartSTR)
            TheHdw.Digital.Patgen.HaltWait
                 
            DSSCCap_MByte_Signal = TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture.Signals("DSSC_Cap_All").DSPWave
            
            ' For debug verify captured segment
            'If AddrIndx = 0 Or AddrIndx = 7 Then
            '    DSSCCap_MByte_Signal.Plot "Read Data Result " & AddrIndx
            '    SPI_Bin_MByte = SPI_Binary_Signal.Select(DSSCCapSize * AddrIndx, 1, DSSCCapSize)
            'End If
    
            If AddrIndx = 0 Then
                DspWaveCap = DSSCCap_MByte_Signal.Copy
            Else
                DspWaveCap = DspWaveCap.Concatenate(DSSCCap_MByte_Signal)
            End If
               
            AddrIndx = AddrIndx + 1
            
        Loop
    
    Close #1
    
    '==========================================================
    
    TheHdw.Wait 0.05
    
    Dim Different_Value As New DSPWave
    Dim Abs_Different_Value As New DSPWave
    Dim IndexOfMaximumValue As Long
    Dim IndexOfMinimumValue As Long
    Dim MaximumValue As Double
    Dim MinimumValue As Double
    Dim SPI_Flag As New SiteLong
    Dim site As Variant
    Dim Sum As New DSPWave
    
    For Each site In TheExec.sites
        Sum = DspWaveCap
        Different_Value = SPI_Binary_Signal.Subtract(Sum)
        Abs_Different_Value = Different_Value.Abs
        MaximumValue = Abs_Different_Value.CalcMaximumValue(IndexOfMaximumValue)
        If MaximumValue = 0 Then
            SPI_Flag = 1        'SPI-ROM PROGRAMMED with current ROM CODE
        Else
            SPI_Flag = 0        'SPI-ROM ERASE
        End If
        TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "Initial Read"
    Next site
        
    '============================================================
        
    Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function



Public Function SPIROM_CheckSum_32MByte(PatName As Pattern, PatName2 As Pattern, DSSCCapPin As PinList, RomSize As Long, RomFileName1 As String) As Long
      
Dim Rom_Data As Byte
Dim TotalDspArray() As Long
Dim DspArray1() As Long
Dim DspArray2() As Long
Dim Count As Long
Dim SrcWaveName As String
Dim StartSTR As String
Dim result As New SiteDouble
Dim WaveCount As Long
Dim SrcDataSize As Long
Dim DSSCCapSize As Long
Dim DspWaveCap As New DSPWave
Dim DSSCCap_MByte_Signal As New DSPWave
Dim DSSCCap_All_Signal_SUM As New DSPWave

Dim SPI_Binary_Signal1 As New DSPWave
Dim SPI_Binary_Signal2 As New DSPWave
Dim SPI_Bin_MByte As New DSPWave


    On Error GoTo ErrorHandler

    TheHdw.Utility.Pins("K3_SPI0_SCLK_SSIN_ROM,K5_SPI0_MISO_MOSI_ROM").State = tlUtilBitOn

    ' Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.5

    If glb_TesterType = "Jaguar" Then TheHdw.Digital.Patterns.Pat(PatName).Load
    TheHdw.Digital.Patgen.TimeOut = 800
   
   
    DSSCCapSize = 4194304
'    SrcDataSize = RomSize
'    Total_Count = 0
'    WaveCount = 0
'
'    ReDim DspArray1(RomSize / 2) As Long
'    ReDim DspArray2(RomSize / 2) As Long
'
'    'ReDim Preserve DspArray(RomSize) As Double
'    For Count = 0 To RomSize - 1
'        If Count <= (RomSize / 2) Then
'            DspArray1(Count) = 2 ^ 8 - 1
'        Else
'            DspArray2(Count - ((RomSize / 2) + 1)) = 2 ^ 8 - 1
'        End If
'    Next Count
'
'    '==========================================================
'
'    Open RomFileName1 For Binary As #1
'
'
'        ' Create data array for reference
'        Count = 0
'        While ((Not EOF(1)) And (Count < SrcDataSize))
'            Get #1, Total_Count + 1, Rom_Data
'            If Count <= (RomSize / 2) Then
'                DspArray1(Count) = Rom_Data
'            Else
'                DspArray2(Count - ((RomSize / 2) + 1)) = Rom_Data
'            End If
'            Count = Count + 1
'            Total_Count = Total_Count + 1
'            'Debug.Print Rom_Data
'        Wend
'        SPI_Binary_Signal1.Data = DspArray1
'        SPI_Binary_Signal2.Data = DspArray2
'        SPI_Binary_Signal = SPI_Binary_Signal1.Concatenate(SPI_Binary_Signal2)
                
                
    Dim AddrIndx As Integer
    Dim AddrSize As Integer
    
    DspWaveCap.Clear
    
    ' If memory size is 32MByte, pattern read address inc is 1MByte
    ' therefore the total iteration: AddrSize= 32
    AddrSize = 8
       
    AddrIndx = 0
    Do While AddrIndx <> AddrSize
    
        DSSCCap_MByte_Signal.Clear
        With TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture
            .Signals.Add "DSSC_Cap_All"
            If glb_TesterType = "Jaguar" Then .Signals("DSSC_Cap_All").Offset = 0
            .Signals.Item("DSSC_Cap_All").SampleSize = DSSCCapSize
            .Signals("DSSC_Cap_All").LoadSettings
        End With
        TheHdw.Wait 0.05
    
        StartSTR = "M32Mbyte_4MegCap_Addr" & CStr(AddrIndx)
        TheHdw.Patterns(PatName).start (StartSTR)
        TheHdw.Digital.Patgen.HaltWait
             
        DSSCCap_MByte_Signal = TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture.Signals("DSSC_Cap_All").DSPWave
        
        ' For debug verify captured segment
        'If AddrIndx = 0 Or AddrIndx = 7 Then
        '    DSSCCap_MByte_Signal.Plot "DSP-MByte Checksum Result " & AddrIndx
        '    SPI_Bin_MByte = SPI_Binary_Signal.Select(DSSCCapSize * AddrIndx, 1, DSSCCapSize)
        '    SPI_Bin_MByte.Plot "BIN-MByte Checksum Result " & AddrIndx
        'End If

        If AddrIndx = 0 Then
            DspWaveCap = DSSCCap_MByte_Signal.Copy
        Else
            DspWaveCap = DspWaveCap.Concatenate(DSSCCap_MByte_Signal)
        End If
           
        AddrIndx = AddrIndx + 1
        
    Loop
    
'    Close #1
    
    '==========================================================
    
    TheHdw.Wait 0.05
    
    Dim Different_Value As New DSPWave
    Dim Abs_Different_Value As New DSPWave
    Dim IndexOfMaximumValue As Long
    Dim IndexOfMinimumValue As Long
    Dim MaximumValue As Double
    Dim MinimumValue As Double
    Dim SPI_Flag As New SiteLong
    Dim site As Variant
    Dim Sum As New DSPWave
    
    For Each site In TheExec.sites
        Sum = DspWaveCap
        Different_Value = SPI_Binary_Signal.Subtract(Sum)
        Abs_Different_Value = Different_Value.Abs
        MaximumValue = Abs_Different_Value.CalcMaximumValue(IndexOfMaximumValue)
        If MaximumValue = 0 Then
            SPI_Flag = 1
        Else
            SPI_Flag = 0
        End If
        TheExec.Flow.TestLimit SPI_Flag, 1, 1, , , , , , "Checksum Result"
    Next site
        
    '============================================================
        
    Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function



Public Function SPIROM_Read_VBT(PatName As Pattern, DSSCCapPin As PinList) As Long
 
On Error GoTo ErrorHandler
                           
    Dim DSSCCap_All_Signal As New DSPWave
    ' Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
              
'     With TheHdw.DCVI("VDDSPIROM30")
'        '.Gate = False
'        .Connect tlDCVIConnectDefault
'        .Mode = tlDCVIModeVoltage
'        .VoltageRange.Autorange = True
'        .NominalBandwidth = 0
'        .Voltage = Power
'        .Current = 0.2
'        .Gate = True
'    End With
    
    TheHdw.Wait 0.5
              
    If glb_TesterType = "Jaguar" Then TheHdw.Digital.Patterns.Pat(PatName).Load
    TheHdw.Digital.Patgen.TimeOut = 800
    
     With TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture
        .Signals.Add "DSSC_Cap_All"
        If glb_TesterType = "Jaguar" Then .Signals("DSSC_Cap_All").Offset = 0
        .Signals.Item("DSSC_Cap_All").SampleSize = 5000000
        .Signals("DSSC_Cap_All").LoadSettings
    End With

       
    TheHdw.Wait 0.05
       
    TheHdw.Patterns(PatName).Test pfAlways, 0

    TheHdw.Digital.Patgen.HaltWait
         
'''    DSSCCap_All_Signal = TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture.Signals("DSSC_Cap_All").DSPWave
'''
'''    DSSCCap_All_Signal.Plot "Read Data Result"
      
    Exit Function
ErrorHandler:
    If AbortTest Then Exit Function Else Resume Next
    
End Function



Public Function SPIROM_Read_MByte_VBT(PatName As Pattern, DSSCCapPin As PinList, AddrIndx As Integer) As Long
Dim MemSize As Variant
Dim DSSCCapSize As Long
Dim DecToBin As String
Dim DspWaveCap As New DSPWave
Dim StartSTR As String
Dim DSSCCap_All_Signal As New DSPWave
Dim DataSize As Long

Dim Wave_Cap_MByte As New DSPWave
Dim SPI_Bin_MByte As New DSPWave
Dim RSLTMByte As Double
    
   
    DSSCCapSize = 4194304
    DataSize = 8
    
     
    On Error GoTo ErrorHandler
                           
    ' Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.5
              
    If glb_TesterType = "Jaguar" Then TheHdw.Digital.Patterns.Pat(PatName).Load
    TheHdw.Digital.Patgen.TimeOut = 800
    
    
    AddrIndx = 0
    Do While AddrIndx <> DataSize
    
        DSSCCap_All_Signal.Clear
        With TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture
            .Signals.Add "DSSC_Cap_All"
            If glb_TesterType = "Jaguar" Then .Signals("DSSC_Cap_All").Offset = 0
            .Signals.Item("DSSC_Cap_All").SampleSize = DSSCCapSize + 1
            .Signals("DSSC_Cap_All").LoadSettings
        End With
        TheHdw.Wait 0.05
    
        StartSTR = "M32Mbyte_4MegCap_Addr" & CStr(AddrIndx)
        TheHdw.Patterns(PatName).start (StartSTR)
        TheHdw.Digital.Patgen.HaltWait
             
        DSSCCap_All_Signal = TheHdw.DSSC.Pins(DSSCCapPin).Pattern(PatName).Capture.Signals("DSSC_Cap_All").DSPWave
        
        ' For Debug
        'If AddrIndx = 0 Or AddrIndx = 31 Then
        '    DSSCCap_All_Signal.Plot "Read Data Result " & AddrIndx
        '    SPI_Bin_MByte = SPI_Binary_Signal.Select(DSSCCapSize * AddrIndx, 1, DSSCCapSize)
        '    rundsp.CompareData SPI_Bin_MByte, DSSCCap_All_Signal, RSLTMByte
        'End If
               
        AddrIndx = AddrIndx + 1
        
    Loop
      
    Exit Function
ErrorHandler:
    If AbortTest Then Exit Function Else Resume Next
    
End Function
Public Function SPIROM_p2p_short_Power(PowerPins As PinList, ForceV As Double) As Long
'
   ''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Testing Method:  Force 0.1V , measure smaller than 199ma,set clamp to 200ma, if higher than 199 ma then fail
    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim HexVSMeasure As New PinListData
    Dim hdvsMeasure As New PinListData
    Dim PowerMeasure As New PinListData
'    Dim PowerPins As String
    Dim p As Variant, Pin_Ary() As String, p_cnt As Long
'    PowerPins = AllHexVsPins & "," & AllUVSPins
       
'    On Error GoTo errHandler
    
'    TheHdw.digital.ApplyLevelsTiming False, True, False, tlPowered 'SEC DRAM
'    pwr_on_i_meter_DCVS AllHexVsPins.Value, 0#, 0.05, 0.01, 0.002, 10, 0.002  'set Force voltage and Current/Meter Range

'    No not need to add the code to avoid SPIROM power pin alarm 20170209
''    If TheExec.EnableWord("HardIP_Alarm_off") = True Then
''    '' 20160419 - Debug Alarm off
''    TheHdw.DCVS.Pins("spi_1v8").Alarm(tlDCVSAlarmAll) = tlAlarmOff
''    End If
    DCVS_PowerOn_I_Meter PowerPins.Value, 0#, 200 * uA, 0.01, 0.002, 10, 0.002
    TheHdw.Wait 0.01
    TheExec.DataManager.DecomposePinList PowerPins, Pin_Ary, p_cnt
    
    For Each p In Pin_Ary
        TheHdw.DCVS.Pins(p).Voltage.Main.Value = ForceV
        TheHdw.Wait 0.05
        DCVS_MeterRead DCVS_UVS256, CStr(p), 10, PowerMeasure
        
        'offline mode simulation  20160328
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                PowerMeasure.Pins(p).Value(site) = -0.005 + Rnd() * 0.0001
            Next site
        End If
        
        
        TheExec.Flow.TestLimit resultVal:=PowerMeasure, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
        
        If TheExec.sites.Active.Count = 0 Then Exit Function 'chihome
        TheHdw.DCVS.Pins(p).Voltage.Main.Value = 0#
    Next p
    Exit Function
 
errHandler:
 
 TheExec.Datalog.WriteComment " p2p_short_Power_SPIROM."
    
End Function
Public Function SPIROM_WaitTime(argc As Long, argv() As String) As Long
    TheHdw.Wait 0.5 'interpose function
End Function


Public Function SPIROM_SectorErase(PatName As Pattern, DSSCSrcPin As PinList, RelayOnPins As PinList) As Long
    
    On Error GoTo ErrorHandler
    Dim DspSrcWave As New DSPWave
    Dim i, j, k As Integer
    Dim w_Addr(31) As Long
    Dim w_4MSize As Long
    Dim w_4MSize_idx As Long
    Dim PatCount As Long
    Dim PattArray() As String
    Dim patt As String
    
    TheHdw.Utility.Pins(RelayOnPins).State = tlUtilBitOn
    ' Level & Timing
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call PATT_GetPatListFromPatternSet(PatName.Value, PattArray, PatCount)
    patt = CStr(PattArray(0))
    If glb_TesterType = "Jaguar" Then TheHdw.Digital.Patterns.Pat(patt).Load
    TheHdw.Digital.Patgen.TimeOut = 10
    
    
    For w_4MSize_idx = 0 To 4194304 Step 65536
'    For w_4MSize_idx = 0 To 65536 Step 65536
'    For w_4MSize_idx = 0 To 0 Step 1
    
        
        w_4MSize = w_4MSize_idx
        For k = 0 To 31
            w_Addr(k) = 0
        Next k
        i = 31
        Do While w_4MSize > 0
            w_Addr(i) = w_4MSize Mod 2
            w_4MSize = w_4MSize \ 2
            i = i - 1
        Loop
        'WaveDefinitions
        DspSrcWave.Data = w_Addr
'        DspSrcWave.Plot
        Call TheExec.WaveDefinitions.CreateWaveDefinition("SrcData", DspSrcWave, True)

        'Insert DSSC Loading & Pattern Loading & Pattern Execution.
        With TheHdw.DSSC.Pins(DSSCSrcPin).Pattern(patt).Source
            .Signals.Add "DSSC_SRC_Addr"
            .Signals.DefaultSignal = "DSSC_SRC_Addr"
            .Signals("DSSC_SRC_Addr").WaveDefinitionName = "SrcData"
            .Signals.Item("DSSC_SRC_Addr").SampleSize = 32
            .Signals("DSSC_SRC_Addr").Amplitude = 1
            If glb_TesterType = "Jaguar" Then .Signals("DSSC_SRC_Addr").LoadSettings
        End With

        TheHdw.Patterns(patt).Test pfNever, 0
        TheHdw.Digital.Patgen.HaltWait
        TheHdw.Wait 0.1
    Next w_4MSize_idx
    
    
Exit Function
    
ErrorHandler:
     If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function Print_Footer_SPIROM(PrintInfo As String)

    If TheExec.DataManager.instanceName Like "SPIROM_Footer_2" Then
        Call SPI_ROM_Written_record
    End If
    TheExec.Datalog.WriteComment "******************************"
    TheExec.Datalog.WriteComment "*print: " & PrintInfo & " end*"
    TheExec.Datalog.WriteComment "******************************"

End Function

Public Function SPI_ROM_writtten_record()

        Call SPI_ROM_Written_record
    
End Function

Public Function Print_Header_SPIROM(PrintInfo As String)

    TheExec.Datalog.WriteComment "********************************"
    TheExec.Datalog.WriteComment "*print: " & PrintInfo & " start*"
    TheExec.Datalog.WriteComment "********************************"

End Function
