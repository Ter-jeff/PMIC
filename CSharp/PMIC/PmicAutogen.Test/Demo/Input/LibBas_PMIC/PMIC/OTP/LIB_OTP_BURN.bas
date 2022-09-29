Attribute VB_Name = "LIB_OTP_BURN"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

'20190329 New method, OTP-DSP
'___Read the captured data DSPWave with size of g_iOTP_REGDATA_BW
'Public Function OTP_READREG_DSP(ByVal sPatName As String, g_sTDI As String, g_sTDO As String, ByVal lAddr As Long) ', svData As SiteVariant)
Public Function OTP_READREG_DSP(ByVal sPatName As String, ByVal lAddr As Long)
    Dim sFuncName As String: sFuncName = "OTP_READREG_DSP"
    On Error GoTo ErrHandler
    Dim sSignalName As String: sSignalName = "OnlyAddress"
    Dim sWaveDef As String: sWaveDef = "Wavedef_Read_allSites"
    Dim sDLogStr As String
    Dim wDataBitWave As New DSPWave ''''it's binary data bits
    Dim wDataValWave As New DSPWave ''''it's decimal
    Dim svDSSCCapBinStr As New SiteVariant
    Dim svData As New SiteVariant: svData = 0
    Dim iBitIndex As Integer
    ''''because addrVal is same in all sites
    Dim wAddrValWave As New DSPWave
    Dim wAddrBitWave As New DSPWave
    Dim adAddrBit() As Double '''<MUST> be Double array for CreateWaveDefinition
    
    
    If TheExec.Sites.Active.Count = 0 Then Exit Function

    wAddrValWave.CreateConstant (lAddr + g_iOTP_ADDR_OFFSET), 1, DspLong
    
    If TheExec.TesterMode = testModeOnline Then
        '___Online test
        'setup capture
        ''''The below action:
        ''''1. setup DigCapture and Execute the pattern
        ''''2. get the captured data DSPWave
        TheHdw.Patterns(sPatName).Load
        For Each Site In TheExec.Sites
            wAddrBitWave = wAddrValWave.ConvertStreamTo(tldspSerial, g_iOTP_ADDR_BW, 0, Bit0IsMsb)
            adAddrBit = wAddrBitWave.Data
            ''''only do one site to get adAddrBit()
            Exit For ''''20200313 enable
        Next Site
        
        TheExec.WaveDefinitions.CreateWaveDefinition sWaveDef, adAddrBit, True
        TheHdw.DSSC.Pins(g_sTDI).Pattern(sPatName).Source.Signals.Add sSignalName
        With TheHdw.DSSC(g_sTDI).Pattern(sPatName).Source.Signals(sSignalName)
               .WaveDefinitionName = sWaveDef
               .Amplitude = 1
               .SampleSize = g_iOTP_ADDR_BW
               '.LoadSamples
               .LoadSettings
        End With
    
        TheHdw.DSSC(g_sTDI).Pattern(sPatName).Source.Signals.DefaultSignal = sSignalName
        Call DSSC_Capture_Setup_OTP(sPatName, "dataSig", g_iOTP_REGDATA_BW, wDataBitWave)
    Else
        '___Offline simulation
        For Each Site In TheExec.Sites
            wDataBitWave = gD_wPGMData.Select(lAddr * gD_slOTP_REGDATA_BW, 1, gD_slOTP_REGDATA_BW).Copy
        Next Site

    End If
    
       
    ''''<TRICK and NOTICE>
    ''''Here using Site loop to make sure capWave is Ready when it's Automatic Mode
    For Each Site In TheExec.Sites
        If (wDataBitWave.SampleSize <> gD_slOTP_REGDATA_BW) Then GoTo ErrHandler
    Next Site
    '___Allocate the wDataBitWave to the proper location of whole read dspwave
    Call RunDsp.otp_get_read_DataWave(wDataBitWave, lAddr, wDataValWave, svData)

            
    '___B). Debug only:
    If (g_bOTPDsscBitsDebugPrint) Then
        For Each Site In TheExec.Sites
            sDLogStr = ""
            For iBitIndex = gD_slOTP_REGDATA_BW - 1 To 0 Step -1 'do string reverse here
                svDSSCCapBinStr = svDSSCCapBinStr & wDataBitWave.Element(iBitIndex)
            Next iBitIndex
        
            sDLogStr = sDLogStr & "Addr:" & FormatLog(lAddr, 4) & ", Site(" & _
                        CStr(Site) & "):  DSSC Capture[Data LSB(Bit0)-MSB)] = " & _
                        " LSB [" & svDSSCCapBinStr(Site) & "] MSB(Bit0)"
            sDLogStr = sDLogStr & ",  Decimal=" & CStr(ConvertFormat_Bin2Dec(CStr(svDSSCCapBinStr(Site))))
            TheExec.Datalog.WriteComment sDLogStr
         Next Site
    End If

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___Write OTP data (size g_iOTP_REGDATA_BW * g_iOTP_ADDR_TOTAL)  once with DigSource Pattern.
Public Function OTP_WRITEREG_DSP_ALL(r_sPatName As String, r_MaskECIDDataWave As DSPWave)
    Dim sFuncName As String: sFuncName = "OTP_WRITEREG_DSP_ALL"
    On Error GoTo ErrHandler
    
    Dim lDataWidth As Long
    Dim lSrcSampleSize As Long
    Dim wAllSrcData As New DSPWave
    Dim sSignalName As String: sSignalName = "AllSourceData"
    Dim sWaveDef As String: sWaveDef = "WaveDef"

    If TheExec.Sites.Active.Count = 0 Then Exit Function
    
    'Remove (2018/05/15)
    'Address = ToSiteLong(Addr)
    'Address = Address.BitwiseOr(g_iOTP_ADDR_OFFSET)
    
    lDataWidth = g_iOTP_REGDATA_BW   '32 * 1= 32
    lSrcSampleSize = g_iOTP_REGDATA_BW * CLng(g_iOTP_ADDR_TOTAL) '32*1024
    
    If TheExec.Sites.Active.Count = 0 Then Exit Function
     ' Build DSPWave
    If TheExec.TesterMode = testModeOnline Then
        

        '___20200313, need to check
        ''''MP7P uses ECID_Mask() outside
        For Each Site In TheExec.Sites
            If g_sbOtpedECID = True Then
                wAllSrcData = r_MaskECIDDataWave.Copy
            Else
                wAllSrcData = gD_wPGMData.Copy
            End If
        Next Site
        
        '==CreateDSSCSource:
        TheHdw.Patterns(r_sPatName).Load
        
        For Each Site In TheExec.Sites
            'wAllSrcData = gD_wPGMData.Copy '7P update gD_wPGMData outsite
            TheExec.WaveDefinitions.CreateWaveDefinition sWaveDef & Site, wAllSrcData, True
            TheHdw.DSSC.Pins(g_sTDI).Pattern(r_sPatName).Source.Signals.Add sSignalName
            With TheHdw.DSSC.Pins(g_sTDI).Pattern(r_sPatName).Source.Signals(sSignalName)
                .WaveDefinitionName = sWaveDef & Site
                .SampleSize = wAllSrcData.SampleSize
                .Amplitude = 1
                '.LoadSamples
                .LoadSettings
            End With
            If (False) Then
                wAllSrcData.Plot
            End If
        Next Site
        
        TheHdw.DSSC.Pins(g_sTDI).Pattern(r_sPatName).Source.Signals.DefaultSignal = sSignalName
        TheHdw.Patterns(r_sPatName).Start ("")
        
        TheHdw.Digital.Patgen.Continue 0, cpuA                              'clear flag
        ' Bypass DSP computing, use HOST computer
        TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
        ' Halt on opcode to make sure all samples are capture.
        TheHdw.Digital.Patgen.HaltWait
    End If
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



''''20190329 New method, OTP-DSP
'___Read OTP data (size g_iOTP_REGDATA_BW * g_iOTP_ADDR_TOTAL)  once with DigCap Pattern.
'Public Function OTP_READREG_DSP_ALL(ByVal sPatName As String, g_sTDI As String, g_sTDO As String)
Public Function OTP_READREG_DSP_ALL(ByVal sPatName As String)
    Dim sFuncName As String: sFuncName = "OTP_READREG_DSP_ALL"
    On Error GoTo ErrHandler
    Dim lDataWidth As Long
    Dim iBitIndex As Long
    Dim wDataBitWave As New DSPWave ''''it's binary data bits
    Dim svDSSCCapBinStr As New SiteVariant
    Dim sDLogStr As String
    Dim wDataBitWaveBySec As New DSPWave
    Dim lOtpAddrIdex As Long

    If TheExec.Sites.Active.Count = 0 Then Exit Function

    lDataWidth = g_iOTP_REGDATA_BW * CLng(g_iOTP_ADDR_TOTAL)   ' 32 * 512
    wDataBitWave.CreateConstant 0, lDataWidth

    If TheExec.TesterMode = testModeOffline Then
        '___Offline simulation
        For Each Site In TheExec.Sites
            ''wDataBitWave = gD_wPGMData.Select(0 * gD_slOTP_REGDATA_BW, 1, lDataWidth).Copy
            ''gD_wReadData = wDataBitWave
            gD_wReadData = gD_wPGMData.Copy ''''20200313 update
        Next Site
    Else
        TheHdw.Patterns(sPatName).Load
    
        '___Online test
        'setup capture
        ''''The below action:
        ''''1. setup DigCapture and Execute the pattern
        ''''2. get the captured data DSPWave
        'TheHdw.DSSC(g_sTDI).Pattern(sPatName).Source.Signals.DefaultSignal = SignalName
        Call DSSC_Capture_Setup_OTP(sPatName, "dataSig", lDataWidth, wDataBitWave)
        
        ''''<TRICK and NOTICE> 20200313 add
        ''''Here using Site loop to make sure capWave is Ready when it's Automatic Mode
        For Each Site In TheExec.Sites
            If (wDataBitWave.SampleSize <> lDataWidth) Then GoTo ErrHandler
            ''''20200313 could be as below because of the one-shot, no need to select
            gD_wReadData = wDataBitWave.Copy
        Next Site
        
    End If
    

    '___B) Debug only:
    If (g_bOTPDsscBitsDebugPrint) Then
         For Each Site In TheExec.Sites
            For lOtpAddrIdex = 0 To g_iOTP_ADDR_TOTAL - 1
                svDSSCCapBinStr = "" '20190517
                
                wDataBitWaveBySec = wDataBitWave.Select(lOtpAddrIdex * gD_slOTP_REGDATA_BW, 1, gD_slOTP_REGDATA_BW).Copy
                
                 For iBitIndex = gD_slOTP_REGDATA_BW - 1 To 0 Step -1 'do string reverse here
                     svDSSCCapBinStr = svDSSCCapBinStr & wDataBitWaveBySec.ElementLite(iBitIndex)
                 Next iBitIndex
                
                 sDLogStr = "Addr:" & FormatLog(lOtpAddrIdex, 4) & ", Site(" & _
                             CStr(Site) & "):  DSSC Capture[Data LSB(Bit0)-MSB)] = " & _
                             " MSB [" & svDSSCCapBinStr(Site) & "] LSB(Bit0)"
                 sDLogStr = sDLogStr + ",  Decimal=" & CStr(ConvertFormat_Bin2Dec(CStr(svDSSCCapBinStr(Site))))
                 TheExec.Datalog.WriteComment sDLogStr
             
            Next lOtpAddrIdex
          Next Site
    End If
    

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'******************************************************************************
' Digital Signal Capture utilities
'******************************************************************************
Public Function DSSC_Capture_Setup_OTP(r_sPatName As String, r_sSignalName As String, _
                                    r_lSampleSize As Long, r_wCapWave As DSPWave)
    Dim sFuncName As String: sFuncName = "DSSC_Capture_Setup_OTP"
    On Error GoTo ErrHandler
 
    TheHdw.Patterns(r_sPatName).Load
    With TheHdw.DSSC.Pins(g_sTDO).Pattern(r_sPatName).Capture.Signals
        .Add (r_sSignalName)
        With .Item(r_sSignalName)
            .SampleSize = r_lSampleSize
            .LoadSettings
        End With
    End With
           
    TheHdw.Patterns(r_sPatName).Start ("")
    
    TheHdw.Digital.Patgen.Continue 0, cpuA                              'clear flag
    TheHdw.Digital.Patgen.FlagWait 0, cpuA                              ' Wait for completion
    TheHdw.Digital.Patgen.Continue 0, cpuA                              'clear flag

    TheHdw.Digital.Patgen.HaltWait

    r_wCapWave = TheHdw.DSSC.Pins(g_sTDO).Pattern(r_sPatName).Capture.Signals(r_sSignalName).DSPWave
    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
       
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'******************************************************************************
' Digital Signal Source utilities
'******************************************************************************
''''==========================================================================================================
'''''EXAMPLE: (Add:2017/10/31)
'''''Setup Source:
''''    Dim  SrcDataArray As New DSPWave
''''    '....put the code to build DSPWave.
''''    Call DSSC_Source_Setup_OTP(r_sPatName, DigSrc_PinS, "Address_data", DigSrc_SampleSize, SrcDataArray)
'''''Setup capture & RUN
''''    Dim  DataArray As New DSPWave
''''    Call DSSC_Capture_Setup_OTP(r_sPatName, "dataSig", DigCap_SampleSize, DataArray)
'''''==========================================================================================================
Public Function DSSC_Source_Setup_OTP(r_sPatName As String, r_sDigSrcPin As String, _
                r_sSignalName As String, r_lSampleSize As Long, r_wSrcWave As DSPWave)
    Dim sFuncName As String: sFuncName = "DSSC_Source_Setup_OTP"
    On Error GoTo ErrHandler
    Dim sWaveDef As String: sWaveDef = "Wavedef"
    TheHdw.Patterns(r_sPatName).Load

    For Each Site In TheExec.Sites
        TheExec.WaveDefinitions.CreateWaveDefinition sWaveDef & Site, r_wSrcWave, True
        TheHdw.DSSC.Pins(r_sDigSrcPin).Pattern(r_sPatName).Source.Signals.Add r_sSignalName
        With TheHdw.DSSC(r_sDigSrcPin).Pattern(r_sPatName).Source.Signals(r_sSignalName)
            .WaveDefinitionName = sWaveDef & Site
            .Amplitude = 1
            .SampleSize = r_lSampleSize
            '.LoadSamples ''''20200313
            .LoadSettings
        End With
        If (False) Then 'DEBUG USED
            r_wSrcWave.Plot
            Stop
        End If
    Next Site
    TheHdw.DSSC(r_sDigSrcPin).Pattern(r_sPatName).Source.Signals.DefaultSignal = r_sSignalName


Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'___OTP Burn Pattern init
Public Function OTP_WRITEREG_initPatt(r_psJTG_Write_Setup As PatternSet)
    Dim sFuncName As String: sFuncName = "OTP_WRITEREG_initPatt"
    On Error GoTo ErrHandler

    TheExec.Datalog.WriteComment ("JTG_EFUSE WRITE SETUP")
    TheExec.Datalog.WriteComment ("RUN PAT:" & r_psJTG_Write_Setup.Value)
    TheHdw.Patterns(r_psJTG_Write_Setup).Load
    TheHdw.Patterns(r_psJTG_Write_Setup).Start ("")
    TheHdw.Digital.Patgen.Continue 0, cpuA                              'clear flag
  
    TheHdw.Digital.Patgen.HaltWait

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'___OTP-DSP 20190327
'___Write the OTP data in multi-shot mode, with the size of (lAddrDataWidth = g_iOTP_ADDR_BW + g_iOTP_REGDATA_BW)
'Public Function OTP_WRITEREG_DSP(r_sPatName As String, g_sTDI As String, g_sTDO As String, ByVal lAddr As Long, r_wAddrDataWave As DSPWave)
Public Function OTP_WRITEREG_DSP(r_sPatName As String, ByVal lAddr As Long, r_wAddrDataWave As DSPWave)
    Dim sFuncName As String: sFuncName = "OTP_WRITEREG_DSP"
    On Error GoTo ErrHandler
    Dim sSignalName As String: sSignalName = "Addressplusdata"
    Dim sWaveDef As String: sWaveDef = "WaveDef"
    Dim lAddrDataWidth As Long
    
    lAddrDataWidth = g_iOTP_ADDR_BW + g_iOTP_REGDATA_BW
    If TheExec.Sites.Active.Count = 0 Then Exit Function
    If TheExec.TesterMode = testModeOnline Then
    
        TheHdw.Patterns(r_sPatName).Load
        For Each Site In TheExec.Sites
            TheExec.WaveDefinitions.CreateWaveDefinition sWaveDef & Site, r_wAddrDataWave, True
            TheHdw.DSSC.Pins(g_sTDI).Pattern(r_sPatName).Source.Signals.Add sSignalName
            With TheHdw.DSSC.Pins(g_sTDI).Pattern(r_sPatName).Source.Signals(sSignalName)
                .WaveDefinitionName = sWaveDef & Site
                .SampleSize = lAddrDataWidth
                .Amplitude = 1
                '.LoadSamples
                .LoadSettings
            End With
        Next Site
    
        TheHdw.DSSC.Pins(g_sTDI).Pattern(r_sPatName).Source.Signals.DefaultSignal = sSignalName
        TheHdw.Patterns(r_sPatName).Start ("")
        TheHdw.Digital.Patgen.Continue 0, cpuA
        TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
        TheHdw.Digital.Patgen.HaltWait
    End If
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



'___OTP-DSP 20191008
'___Write the OTP data in multi-shot mode, with the size of (lAddrDataWidth = g_iOTP_ADDR_BW + g_iOTP_REGDATA_BW)
Public Function OTP_WRITEREG_DSP_LoopAddr(r_sPatName As String, ByVal laddrStart As Long, ByVal laddrEnd As Long, ByVal wAllAddrDataBitWave As DSPWave)
    Dim sFuncName As String: sFuncName = "OTP_WRITEREG_DSP_LoopAddr"
    On Error GoTo ErrHandler

    Dim lBitsidx As Long
    Dim lAddrIdx As Long
    Dim sSignalName As String: sSignalName = "Addressplusdata"
    Dim sWaveDef As String: sWaveDef = "WaveDef"
    Dim sTempStringBit0IsMsb As String
    Dim sDLogStr As String
    Dim lAddrDataWidth As Long
    Dim AddrDataBitWave As New DSPWave
    Dim lSelectedStartBit As Long

    lAddrDataWidth = g_iOTP_ADDR_BW + g_iOTP_REGDATA_BW
    
    If TheExec.Sites.Active.Count = 0 Then Exit Function
    
    If TheExec.TesterMode = testModeOnline Then
        TheHdw.Patterns(r_sPatName).Load
    
        TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug

        TheHdw.DSSC.Pins(g_sTDI).Pattern(r_sPatName).Source.Signals.Add sSignalName
        For lAddrIdx = laddrStart To laddrEnd
            lSelectedStartBit = lAddrIdx * lAddrDataWidth
            
            For Each Site In TheExec.Sites
                ''''get the related addr+data bits per addr
                AddrDataBitWave = wAllAddrDataBitWave.Select(lSelectedStartBit, 1, lAddrDataWidth).Copy
            
                TheExec.WaveDefinitions.CreateWaveDefinition sWaveDef & Site, AddrDataBitWave, True
                With TheHdw.DSSC.Pins(g_sTDI).Pattern(r_sPatName).Source.Signals(sSignalName)
                    .WaveDefinitionName = sWaveDef & Site
                    .SampleSize = lAddrDataWidth
                    .Amplitude = 1
                    '.LoadSamples
                    .LoadSettings
                End With
            Next Site
            
            If (g_bOTPDsscBitsDebugPrint) Then
                Dim alTmpData() As Long
                Dim lAddr As Long
                lAddr = lAddrIdx
                For Each Site In TheExec.Sites
                    sTempStringBit0IsMsb = ""
                    alTmpData = AddrDataBitWave.Data
                    For lBitsidx = 0 To UBound(alTmpData)
                        sTempStringBit0IsMsb = sTempStringBit0IsMsb + CStr(alTmpData(lBitsidx))
                    Next lBitsidx
                    
                    sDLogStr = "Addr:" & FormatLog(lAddr, 4) & ", Site(" & CStr(Site) & _
                    "):  DSSC Source[Addr(LSB-MSB)+Data(LSB-MSB)] sTempStringBit0IsMsb= " & _
                    Mid(sTempStringBit0IsMsb, 1, g_iOTP_ADDR_BW) & Space(2) & _
                    Mid(sTempStringBit0IsMsb, g_iOTP_ADDR_BW + 1, Len(sTempStringBit0IsMsb) - 9)
                    
                    TheExec.Datalog.WriteComment sDLogStr
                Next Site
            End If
            
            TheHdw.DSSC.Pins(g_sTDI).Pattern(r_sPatName).Source.Signals.DefaultSignal = sSignalName
            TheHdw.Patterns(r_sPatName).Start ("")
            TheHdw.Digital.Patgen.Continue 0, cpuA
            TheHdw.Digital.Patgen.HaltWait
        Next lAddrIdx
    End If
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191203 New method, OTP-DSP
'___Read the captured data DSPWave with size of g_iOTP_REGDATA_BW
Public Function OTP_READREG_DSP_LoopAddr(r_sPatName As String, ByVal laddrStart As Long, ByVal laddrEnd As Long)
    On Error GoTo ErrHandler
    Dim sFuncName As String:: sFuncName = "OTP_READREG_DSP_LoopAddr"
    Dim lBitsidx As Long
    Dim sSignalName As String
    Dim sWaveDef As String
    Dim wDataBitWave As New DSPWave ''''it's binary data bits
    Dim wConbineWave As New DSPWave
    Dim wDataValWave As New DSPWave ''''it's decimal
    Dim svDSSCCapBinStr As New SiteVariant
    Dim sDLogStr As String
    Dim lAddr As Long
    Dim svData As New SiteVariant: svData = 0
    
    If TheExec.Sites.Active.Count = 0 Then Exit Function

    ''''because addrVal is same in all sites
    Dim wAddrValWave As New DSPWave
    Dim wAddrBitWave As New DSPWave
    Dim adAddrBit() As Double '''<MUST> be Double array for CreateWaveDefinition
    
    TheHdw.Patterns(r_sPatName).Load
    
    For lAddr = laddrStart To laddrEnd
        wAddrValWave.CreateConstant (lAddr + g_iOTP_ADDR_OFFSET), 1, DspLong
        'TheHdw.Patterns(r_sPatName).Load
        For Each Site In TheExec.Sites
            wAddrBitWave = wAddrValWave.ConvertStreamTo(tldspSerial, g_iOTP_ADDR_BW, 0, Bit0IsMsb)
            adAddrBit = wAddrBitWave.Data
            ''''only do one site to get adAddrBit()
            Exit For ''''20200313 enable
        Next Site
        
        sWaveDef = "Wavedef_Read_allSites"
        sSignalName = "OnlyAddress"
        TheExec.WaveDefinitions.CreateWaveDefinition sWaveDef, adAddrBit, True
        TheHdw.DSSC.Pins(g_sTDI).Pattern(r_sPatName).Source.Signals.Add sSignalName
        With TheHdw.DSSC(g_sTDI).Pattern(r_sPatName).Source.Signals(sSignalName)
               .WaveDefinitionName = sWaveDef
               .Amplitude = 1
               .SampleSize = g_iOTP_ADDR_BW
               '.LoadSamples
               .LoadSettings
        End With
        
        If TheExec.TesterMode = testModeOnline Then
            '___Online test
            'setup capture
            ''''The below action:
            ''''1. setup DigCapture and Execute the pattern
            ''''2. get the captured data DSPWave
            TheHdw.DSSC(g_sTDI).Pattern(r_sPatName).Source.Signals.DefaultSignal = sSignalName
            Call DSSC_Capture_Setup_OTP(r_sPatName, "dataSig", g_iOTP_REGDATA_BW, wDataBitWave)
        Else
            '___Offline simulation
            For Each Site In TheExec.Sites
                wDataBitWave = gD_wPGMData.Select(lAddr * gD_slOTP_REGDATA_BW, 1, gD_slOTP_REGDATA_BW).Copy
            Next Site
        End If
      
        'wConbineWave = wConbineWave.Concatenate(wDataBitWave)

        ''''<TRICK and NOTICE>
        ''''Here using Site loop to make sure capWave is Ready when it's Automatic Mode
        For Each Site In TheExec.Sites
            If (wDataBitWave.SampleSize <> gD_slOTP_REGDATA_BW) Then GoTo ErrHandler
        Next Site
        '___Allocate the wDataBitWave to the proper location of whole read dspwave
        Call RunDsp.otp_get_read_DataWave(wDataBitWave, lAddr, wDataValWave, svData)
        'Call RunDsp.otp_get_read_DataWave_LoopAddr(wConbineWave, laddrStart, laddrEnd, wDataValWave)
        
        '___B). Debug only:
        If (g_bOTPDsscBitsDebugPrint) Then
             For Each Site In TheExec.Sites
                sDLogStr = ""
                svDSSCCapBinStr = ""
                For lBitsidx = gD_slOTP_REGDATA_BW - 1 To 0 Step -1 'do string reverse here
                    svDSSCCapBinStr = svDSSCCapBinStr & wDataBitWave.ElementLite(lBitsidx)
                Next lBitsidx
             
                sDLogStr = sDLogStr & "Addr:" & FormatLog(lAddr, 4) & ", Site(" & _
                           CStr(Site) & "):  DSSC Capture[Data LSB(Bit0)-MSB)] = " & _
                           " LSB [" & svDSSCCapBinStr(Site) & "] MSB(Bit0)"
                sDLogStr = sDLogStr & ",  Decimal=" & CStr(ConvertFormat_Bin2Dec(CStr(svDSSCCapBinStr(Site))))
                TheExec.Datalog.WriteComment sDLogStr
            Next Site
        End If
    Next lAddr
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


