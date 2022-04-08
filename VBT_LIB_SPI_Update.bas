Attribute VB_Name = "VBT_LIB_SPI_Update"

Option Explicit

Public gSB_RTOSBootPatResult As New SiteBoolean
Public gSB_RTOSBistPatResult As New SiteBoolean
Public gDSPData_UART As New PinListData
Public G_cmd1 As String
Public G_cmd2 As String
Public G_cmd3 As String
Public GlobalMergeAry() As String
Public TNTEMP As String
Public PinTEMP As String
Public force_val As Double

Public g_RTOS_FirstSetp As Boolean
Public g_LastRTOSPoint As Boolean
Public g_RTOSNwireChar As Boolean
Public g_RTOS2DFirstPoint As Boolean
Public g_RTOSRampStep As Integer
Public g_RTOS_SceVoltage As New PinListData
Public g_TResultForBinCut As New SiteBoolean


Public Function RTOS_Command(Optional Cmd1 As String, Optional Cmd2 As String, Optional Cmd3 As String, Optional Cmd4 As String, _
            Optional Cmd5 As String, Optional Cmd1TimeOut As Double = 0#, Optional Cmd2TimeOut As Double = 0#, Optional Cmd3TimeOut As Double = 0#, _
            Optional Cmd4TimeOut As Double = 0#, Optional Cmd5TimeOut As Double = 0#) As Long

On Error GoTo errHandler
    
    ReDim GlobalMergeAry(TheExec.sites.Existing.Count) 'for txt data collection
      
    Dim CmdList As Variant 'String
    Dim CmdListStatus As New SiteLong
    Dim powerPin As String
    Dim instanceName As String: instanceName = TheExec.DataManager.instanceName
    Dim CMDTotalTT As Double

      
    If Cmd1 <> "" Then CmdList = Cmd1
    If Cmd2 <> "" Then CmdList = CmdList + Cmd2
    If Cmd3 <> "" Then CmdList = CmdList + Cmd3
    If Cmd4 <> "" Then CmdList = CmdList + Cmd4
    If Cmd5 <> "" Then CmdList = CmdList + Cmd5
    CMDTotalTT = Cmd1TimeOut + Cmd2TimeOut + Cmd3TimeOut + Cmd4TimeOut + Cmd5TimeOut

    'Scenario Run Conditions
    CmdListStatus = 0
    
    TheExec.Datalog.DatalogSuspended = False
    
    If CmdList <> "" Then Set CmdListStatus = SendCmd(CmdList, CMDTotalTT, False)

    TheExec.Flow.TestLimit CmdListStatus, 1, 1       ', , , , , , TestName
    
    RTOS_UART_Print instanceName, CmdListStatus
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function



Public Function RTOS_eFuse_Read(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, dspwavesize As Long, Optional Efuse_Read_Dec_Flag As Boolean = False, Optional Dict_Store_Dec_Name As String = "", _
                                Optional Calc_code As String = "") As Long

    ' Parameter : eFuse Block , eFuse Variable , data , Data Width
    ' Create dictionary , if exist then remove and re-create
    ' MUST :  if necessary , we can set limit if read out value = 0 then bin out .
    
    Dim site As Variant
    Dim Read_Code As New DSPWave
    Dim Read_Value As New DSPWave
    Dim Efuse_Value As New SiteLong
    Dim TempVal As Long
    Dim Efuse_Value_Chk As New SiteVariant
    Dim i As Long
        
    On Error GoTo errHandler
    
    Read_Code.CreateConstant 0, dspwavesize

    If Efuse_Read_Dec_Flag = True Then
        Read_Value.CreateConstant 0, 1
    End If

    For Each site In TheExec.sites

        Efuse_Value(site) = auto_eFuse_GetReadDecimal(FuseType, m_catename, True)
'''''        Efuse_Value(Site) = CLng(Site) + 8
'''----------cal get fused code
        If Calc_code <> "" Then
        'Calc_code = "minus,100"
            If Split(Calc_code, ",")(0) = "minus" Then
                Efuse_Value = Efuse_Value.Subtract(Split(Calc_code, ",")(1))
            End If
        End If
'''----------cal get fused code
        If Efuse_Read_Dec_Flag = True Then
            Read_Value.Element(0) = Efuse_Value(site)
        End If

        TempVal = Efuse_Value(site)
        For i = 0 To dspwavesize - 1
            Read_Code.Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i
        
        If Efuse_Value(site) = 0 Then                                                'If Read out value = 0 then bin out
            Efuse_Value_Chk(site) = 0
        Else
            Efuse_Value_Chk(site) = 1
        End If
        
    Next site
        
    TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=1, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceNone
        
    Call AddStoredCaptureData(Dict_Store_Code_Name, Read_Code)
    
    If Efuse_Read_Dec_Flag = True Then
        Call AddStoredCaptureData(Dict_Store_Dec_Name, Read_Value)
    End If
    
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Read"
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function RTOS_eFuse_Write(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, Flag_Name As String, Optional Efuse_Binary_Write_Flag As Boolean = False, _
                                Optional Calc_code As String) As Long

    ' Parameter : eFuse Block , eFuse Variable , data
    ' Call auto_eFuse_SetPatTestPass_Flag("CFG", "LPDP_C_RX", TheHdw.Digital.Patgen.PatternBurstPassed(Site))
    ' Call auto_eFuse_SetWriteDecimal("CFG", "LPDP_C_RX", BestCode(Site))

    Dim site As Variant
    Dim RTOS_eFuseData_Dict As New SiteVariant
    Dim Data_Temp As String
    Dim m_value As New SiteVariant
    Dim j As Integer
    Dim Pass_Fail_Flag As New SiteBoolean
    On Error GoTo errHandler
    
    If m_catename = "mtr_fused_t2" Then
        For Each site In TheExec.sites
            m_value = 1
        Next site
    Else
        m_value = GetStoredData(Dict_Store_Code_Name)
    End If

    For Each site In TheExec.sites
        If TheExec.Flow.SiteFlag(site, Flag_Name) = 1 Then
            Pass_Fail_Flag(site) = False
        ElseIf TheExec.Flow.SiteFlag(site, Flag_Name) = 0 Then
            Pass_Fail_Flag(site) = True
        Else
            Pass_Fail_Flag(site) = False
            TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
        End If

        Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catename, Pass_Fail_Flag(site), True)
        Call auto_eFuse_SetWriteDecimal(FuseType, m_catename, m_value(site), True)
        
    Next site
    
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Write"
    If AbortTest Then Exit Function Else Resume Next


End Function

'     If TheExec.Flow.EnableWord("UARTOutPut") = True Then
'         SendCmd Cmd, CmdTimeOut, True
'     Else
'         SendCmd Cmd, CmdTimeOut, False
'     End If
    
' End Function



Public Function RTOS_IDS(Cmd As String, Cmdwait As Double) As SiteLong

On Error GoTo errHandler
    
    ReDim GlobalMergeAry(TheExec.sites.Existing.Count) ''for txt data collection
    
    Dim dspData As New PinListData
    Dim LowPins As String
    Dim HighPins As String
    
    Dim LowToHigh As String
    Dim HighToLow As String

    Dim BootResult As New DSPWave
    Dim i As Long, p As Long
    Dim TResult As New SiteLong
    Dim RTOS_IDS_inst As String
    Dim instanceName As String
    RTOS_IDS_inst = TheExec.DataManager.instanceName
    If Cmdwait < 0.0001 Then
        Cmdwait = 0.1
    End If
    
    TResult = SendCmd(Cmd, Cmdwait)
    
    TheExec.Flow.TestLimit TResult, 1, 1, , , , , , "RTOS_IDS_SC94_Cmd"
    
    RTOS_UART_Print instanceName, TResult
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function


Public Function RTOS_Boot(BootUsingPattern As Boolean, Optional BootPattern As String, Optional UseJTAG As Boolean, Optional shmooing As Boolean, Optional ramp As Boolean = False) As Long

On Error GoTo errHandler

    TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = 70
    TheExec.Datalog.ApplySetup
    
    If Not (shmooing) Then
        ReDim GlobalMergeAry(TheExec.sites.Existing.Count)
    End If
    
    Dim site As Variant
    Dim Relay_Device As String
    Dim Relay_Spirom As String
    Dim dspData As New PinListData
    Dim LowPins As String
    Dim HighPins As String
    
    Dim LowToHigh As String
    Dim HighToLow As String

    Dim BootDSP As New DSPWave
    Dim i As Long, p As Long
    Dim TResult As New SiteLong
    Dim LogTimes As Boolean
    Dim TRef_Before_Pat As Double               '<- Code timing
    Dim TExec_Before_Pat As Double              '<- Execution time
    Dim instanceName As String
    instanceName = TheExec.DataManager.instanceName
    '======================================================
    Relay_Device = "k02,k03"  'update for Tonga ''''' "k01,k03"
    Relay_Spirom = "k01,k04"  ' "k02,k04"
    LowPins = "RTOS_Boot_Low"    'Pin group"
    HighPins = "RTOS_Boot_High"  'Pin group"
    HighToLow = ""
    LowToHigh = "RTOS_Boot_L2H"  'Pin group
    '======================================================
    
'    LogTimes = True
'    If (LogTimes = True) Then
'        TRef_Before_Pat = TheExec.Timer(0)
'    End If
    
'    If ramp = True Then
'       RTOS_Voltage_RampUp
'    Else
    'thehdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
    TheExec.Datalog.WriteComment "Before ALT, VDD_PCPU0:" & TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Value & "V"
       TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'    End If
    TheExec.Datalog.WriteComment "After ALT, VDD_PCPU0:" & TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Value & "V"
    TheHdw.Utility.Pins(Relay_Spirom).State = tlUtilBitOff
    TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOff
    TheHdw.Wait 0.003
        'SPI_ROM_1_8V
    TheHdw.DCVS.Pins("SPI_PWR").Connect                     ' re-cycle SPI-ROM power
    TheHdw.DCVS.Pins("SPI_PWR").Voltage.Main.Value = 0#
    TheHdw.Wait 0.02
    TheHdw.DCVS.Pins("SPI_PWR").Voltage.Main.Value = 1.8
    TheHdw.Wait 0.01
    
'    thehdw.Digital.Pins("xo0,jtag_tck").Disconnect
'    Start_Profile_AutoResolution "SPI_PWR", "I", 0, 0, "RTOS", 1


'''//Follow relay switch for Turks//'''
    TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
    TheHdw.Wait 0.05
    

    With TheHdw.Protocol.ports("UART_TX")
        .TimeOut.Enabled = True
        .TimeOut.Value = 1
        .Enabled = True
        .NWire.MaxHoldUntilTimeout.Value = 0.003        ''//3msec*1500=4.5sec
'''        .NWire.MaxWaitUntilTimeout.Value = 0.005
        .NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
    End With
    
    Set dspData = TheHdw.Protocol.ports("UART_TX").NWire.CMEM.DSPWave
     
    TheHdw.Protocol.ports("UART_TX").Modules("UART_boot").start
    
    TheHdw.Wait 0.002
      
      
    
    If BootUsingPattern Then
       TheHdw.Patterns(BootPattern).Load
    
        If UseJTAG Then
'''            TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
            TheHdw.Wait 0.005
            TheHdw.Patterns(BootPattern).start
            TheHdw.Digital.Patgen.HaltWait
        Else
            TheHdw.Digital.Patgen.Continue 0, cpuA
            TheHdw.Patterns(BootPattern).start
            TheHdw.Digital.Patgen.FlagWait cpuA, 0
            TheHdw.Wait 0.003
'''            TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
            TheHdw.Wait 0.01
            TheHdw.Digital.Patgen.Continue 0, cpuA
            TheHdw.Digital.Patgen.HaltWait
        End If
    Else
    
        'TheHdw.Digital.Pins("OE_N").InitState = chInitLo
        
        TheHdw.Digital.Pins(LowPins).InitState = chInitLo
        TheHdw.Digital.Pins(HighPins).InitState = chInitHi
        TheHdw.Wait 0.01
'''        TheHdw.Digital.Pins("SPI1_MISO").InitState = chInitHi ''001 -> 101   '''//Cebu has
'''        TheHdw.Wait 0.05
    
        TheHdw.Digital.Pins(LowToHigh).InitState = chInitHi
        TheHdw.Wait 0.05
        
'''//Follow relay switch for Cebu//'''
'''        TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
'''        TheHdw.Wait 0.05

''''==== For JTAG Debug ====''''
'''    thehdw.Digital.Pins("jtag_tdi,jtag_tdo,jtag_sel,jtag_trstn,jtag_tms, jtag_tck").Disconnect '<+For JTAG mode?
'''    thehdw.Wait 0.01

    End If
    
'    Plot_Profile "SPI_PWR", "RTOS"

    TheHdw.Protocol.ports("UART_TX").IdleWait
    TheHdw.Protocol.ports("UART_TX").Enabled = False

    '9th, Oct 2019
    'After Discussing with customer,TER-Fred add, for TP checking, bypass dsp process error
    '==================================================================Start
    If TheExec.TesterMode = testModeOffline Then
            Dim CompareArrayOffline(649) As Long
            Dim VarSite As Variant
            Dim dspdata_checking_2 As New DSPWave
            

            For i = 0 To 649 Step 4
                    If i < 650 Then CompareArrayOffline(i) = 65 'A
                    If i + 1 < 650 Then CompareArrayOffline(i + 1) = 84 'T
                    If i + 2 < 650 Then CompareArrayOffline(i + 2) = 69 'E
                    If i + 3 < 650 Then CompareArrayOffline(i + 3) = 62 '>
            Next i
            dspdata_checking_2.Data = CompareArrayOffline

            rundsp.CheckBootStatus dspdata_checking_2, TResult          'Check DSP wave status to determine TResult
    
            For Each VarSite In TheExec.sites.Selected
                BootDSP(VarSite) = dspdata_checking_2(VarSite).Copy
            Next VarSite
    Else

   ''below code are the original vbt
            rundsp.CheckBootStatus dspData, TResult    'Check DSP wave status to determine TResult
            BootDSP = dspData.Copy
    End If
     '==================================================================End
    
    Call LogDUTResponse(BootDSP, TResult) 'Copy DSP wave into an output log
    
    TheHdw.Protocol.ports("UART_TX").TimeOut.Value = 30
       
    
    
    'Boot up Configeration
    'TResult = SendCmd("core up acc;", 0.2)  'Plain boot
    TResult = SendCmd("core up acc;", 0.1)  'Plain boot


    
    
    If Not (shmooing) Then RTOS_UART_Print instanceName, TResult
    If Not (shmooing) Then TheExec.Flow.TestLimit TResult, 1, 1, , , , , , "Boot Status"
    DebugPrintFunc "RTOS_BOOT"
    
    
'    Theexec.Flow.TestLimit TResult, 0, 0, , , , , , "Slave_Up_ACC"
    
'    If TheExec.EnableWord("UARTOut") = True Then 'Enable UART Logs
'        SendCmd "pmgr mode", 0.1, , True
'    Else 'Disable UART Logs
'        SendCmd "pmgr mode", 0.1, , False
'    End If
    

'    SendCmd "bbq on", 0.1   ' send same command to all sites
'    SendCmd cmd, 0.6    ' send unique cmd per site
'    For Each site In TheExec.sites.Active
'        If (LogTimes = True) Then
'            TExec_Before_Pat = TheExec.Timer(TRef_Before_Pat)
'            TheExec.DataLog.WriteComment "ElapsedTime Pat Site (" & site & ")" + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
'        End If
'    Next site
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function RTOS_Boot_MTRSNS(BootUsingPattern As Boolean, Optional BootPattern As String, Optional UseJTAG As Boolean, Optional shmooing As Boolean, Optional ramp As Boolean = False) As Long

On Error GoTo errHandler

    TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = 70
    TheExec.Datalog.ApplySetup
    
    If Not (shmooing) Then
        ReDim GlobalMergeAry(TheExec.sites.Existing.Count)
    End If
    
    Dim site As Variant
    Dim Relay_Device As String
    Dim Relay_Spirom As String
    Dim dspData As New PinListData
    Dim LowPins As String
    Dim HighPins As String
    
    Dim LowToHigh As String
    Dim HighToLow As String

    Dim BootDSP As New DSPWave
    Dim i As Long, p As Long
    Dim TResult As New SiteLong
    Dim LogTimes As Boolean
    Dim TRef_Before_Pat As Double               '<- Code timing
    Dim TExec_Before_Pat As Double              '<- Execution time
    Dim instanceName As String
    instanceName = TheExec.DataManager.instanceName
    '======================================================
    Relay_Device = "k36,k38" 'update for Tonga ''''' "k01,k03"
    Relay_Spirom = "k37,k39" ' "k02,k04"
    LowPins = "RTOS_Boot_Low"    'Pin group"
    HighPins = "RTOS_Boot_High"  'Pin group"
    HighToLow = ""
    LowToHigh = "RTOS_Boot_L2H"  'Pin group
    '======================================================
    
'    LogTimes = True
'    If (LogTimes = True) Then
'        TRef_Before_Pat = TheExec.Timer(0)
'    End If
    
    If ramp = True Then
       RTOS_Voltage_RampUp
    Else
       TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If

    TheHdw.Utility.Pins(Relay_Spirom).State = tlUtilBitOff
    TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOff
    TheHdw.Wait 0.003
        
    TheHdw.DCVS.Pins("SPI_PWR").Connect                     ' re-cycle SPI-ROM power
    TheHdw.DCVS.Pins("SPI_PWR").Voltage.Main.Value = 0#
    TheHdw.Wait 0.02
    TheHdw.DCVS.Pins("SPI_PWR").Voltage.Main.Value = 1.8
    TheHdw.Wait 0.01
    
'    thehdw.Digital.Pins("xo0,jtag_tck").Disconnect
'    Start_Profile_AutoResolution "SPI_PWR", "I", 0, 0, "RTOS", 1


'''//Follow relay switch for Turks//'''
    TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
    TheHdw.Wait 0.05
    

    With TheHdw.Protocol.ports("UART_TX")
        .TimeOut.Enabled = True
        .TimeOut.Value = 2
        .Enabled = True
        .NWire.MaxHoldUntilTimeout.Value = 0.003        ''//3msec*1500=4.5sec
'''        .NWire.MaxWaitUntilTimeout.Value = 0.005
        .NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
    End With
    
    Set dspData = TheHdw.Protocol.ports("UART_TX").NWire.CMEM.DSPWave
     
    TheHdw.Protocol.ports("UART_TX").Modules("UART_boot").start
    
    TheHdw.Wait 0.002
      
      
    
    If BootUsingPattern Then
       TheHdw.Patterns(BootPattern).Load
    
        If UseJTAG Then
'''            TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
            TheHdw.Wait 0.005
            TheHdw.Patterns(BootPattern).start
            TheHdw.Digital.Patgen.HaltWait
        Else
            TheHdw.Digital.Patgen.Continue 0, cpuA
            TheHdw.Patterns(BootPattern).start
            TheHdw.Digital.Patgen.FlagWait cpuA, 0
            TheHdw.Wait 0.003
'''            TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
            TheHdw.Wait 0.01
            TheHdw.Digital.Patgen.Continue 0, cpuA
            TheHdw.Digital.Patgen.HaltWait
        End If
    Else
        TheHdw.Digital.Pins("OE_N").InitState = chInitLo ' foe level shifter
        TheHdw.Digital.Pins(LowPins).InitState = chInitLo
        TheHdw.Digital.Pins(HighPins).InitState = chInitHi
        TheHdw.Wait 0.01
'''        TheHdw.Digital.Pins("SPI1_MISO").InitState = chInitHi ''001 -> 101   '''//Cebu has
'''        TheHdw.Wait 0.05
    
        TheHdw.Digital.Pins(LowToHigh).InitState = chInitHi
        TheHdw.Wait 0.05
        
'''//Follow relay switch for Cebu//'''
'''        TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
'''        TheHdw.Wait 0.05

''''==== For JTAG Debug ====''''
'''    thehdw.Digital.Pins("jtag_tdi,jtag_tdo,jtag_sel,jtag_trstn,jtag_tms, jtag_tck").Disconnect '<+For JTAG mode?
'''    thehdw.Wait 0.01

    End If
    
'    Plot_Profile "SPI_PWR", "RTOS"

    TheHdw.Protocol.ports("UART_TX").IdleWait
    TheHdw.Protocol.ports("UART_TX").Enabled = False


    rundsp.CheckBootStatus dspData, TResult  'Check DSP wave status to determine TResult
    
    BootDSP = dspData.Copy
    
    Call LogDUTResponse(BootDSP, TResult) 'Copy DSP wave into an output log
    
    TheHdw.Protocol.ports("UART_TX").TimeOut.Value = 30
       
    
    'Boot up Configeration
    TResult = SendCmd("fs mount SPI1;pmgr mode 0x2202134;", 0.2)  'Plain boot
    TResult = SendCmd("slave up isp mar_sram.con alloc; slave up ane mar_sram.con alloc", 1)
    ''slave up acc;
    
    
    If Not (shmooing) Then RTOS_UART_Print instanceName, TResult
    If Not (shmooing) Then TheExec.Flow.TestLimit TResult, 1, 1, , , , , , "Boot Status"
    DebugPrintFunc "RTOS_BOOT"
    
    
'    Theexec.Flow.TestLimit TResult, 0, 0, , , , , , "Slave_Up_ACC"
    
'    If TheExec.EnableWord("UARTOut") = True Then 'Enable UART Logs
'        SendCmd "pmgr mode", 0.1, , True
'    Else 'Disable UART Logs
'        SendCmd "pmgr mode", 0.1, , False
'    End If
    

'    SendCmd "bbq on", 0.1   ' send same command to all sites
'    SendCmd cmd, 0.6    ' send unique cmd per site
'    For Each site In TheExec.sites.Active
'        If (LogTimes = True) Then
'            TExec_Before_Pat = TheExec.Timer(TRef_Before_Pat)
'            TheExec.DataLog.WriteComment "ElapsedTime Pat Site (" & site & ")" + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
'        End If
'    Next site
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function RTOS_Boot_CZ(argc As Long, argv() As String) As Long
''ZHHUANGF
On Error GoTo errHandler

    Dim dspData As New PinListData
    Dim LowPins As String
    Dim HighPins As String
    
    Dim LowToHigh As String
    Dim HighToLow As String

    Dim BootResult As New DSPWave
    Dim i As Long, p As Long
    Dim TResult As New SiteLong
    
'    thehdw.DCVS.Pins("vdd_cpu").Voltage.Output = tlDCVSVoltageAlt
    'THEHDW.Digital.ApplyLevelsTiming True, True, True, tlPowered
    g_RTOS_FirstSetp = True
    g_LastRTOSPoint = False
    g_RTOS2DFirstPoint = True
    'RTOS_Boot_Up_fail_Power_Up '200205
    RTOS_Boot False, , , True
'    RTOS_Shmoo_Start = True
    'Boot up Configeration
'    TheExec.Flow.TestLimit TResult, 0, 0, , , , , , "Boot Status"
         
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function

Public Function RTOS_RunMetrology(SensorName As String, VddSenseFreq As String, VddSenseHeat As String, VddSensePreheat As String, _
                                    VddVoltageLevelCNT As String, CmdROT As String, CmdROV As String, CmdTMP As String, CmdFinal As String, _
                                    CmdROT_Timeout As Double, CmdROV_Timeout As Double, CmdTMP_Timeout As Double, CmpFinal_Timeout As Double, _
                                    StartVoltage As Double, EndVoltage As Double, VoltageStepNumber As Double, Optional SelsramBit As String)
On Error GoTo errHandler

     Dim CmdList As String
     Dim CmdListStatus As New SiteLong
     
     Dim powerPin As String
     Dim SupplyVoltage As Long
     Dim TRef_Before_Pat As Double               '<- Code timing
     Dim TExec_Before_Pat As Double              '<- Execution time
     Dim LogTimes As Boolean

    TheHdw.PinLevels.ApplyPower
    Shmoo_Save_core_power_per_site_for_Vbump
    TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain


    ''''////===For Ts5/Ta0/Ta1====////'''
    If UCase(SensorName) Like UCase("*ta*") Then
        SendCmd "pmgr ane on", 0.1
    ElseIf UCase(SensorName) Like UCase("*ts5*") Then
        SendCmd "slave up isp mar_sram.con alloc", 0.1
    End If
    ''''////===For Ts5/Ta0/Ta1====////'''


    'Select Sram Start
    Dim uniquesBit As Boolean, site As Variant
    uniquesBit = IsArray(SelsramBit)
    Dim BinCutSelAry() As String
    BinCutSelAry = Split(SelsramBit, ",")
    
      
    'If UBound(BinCutSelAry) = 5 And SelsramBit <> "" Then
    If UBound(BinCutSelAry) = 4 And SelsramBit <> "" Then 'for Tonga
        For Each site In TheExec.sites.Selected
            BinCutSelAry(site) = Decide_Switching_Bit_RTOS(BinCutSelAry(site), g_ApplyLevelTimingValt, "RTOS")
        Next site
        SendCmd BinCutSelAry, 0.1
    Else
        SendCmd Decide_Switching_Bit_RTOS("SSSS", g_ApplyLevelTimingValt, "RTOS"), 0.1 ' for Tonga
        'SendCmd Decide_Switching_Bit_RTOS("SSSSS", g_ApplyLevelTimingValt, "RTOS"), 0.1
    End If

    '//Change to Valt mode
    TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt '' only change the corepower to alt
    TheHdw.Wait 0.005
    
    Dim OffsetRecord As New SiteLong
    'Delete start, 190628, Leonli
    Dim MTRString_TMPS() As String 'TMPS Per site
    Dim MTRString_ROT() As String 'ROT Per site
    Dim MTRString_ROV() As String 'ROV Per site
    'Delete End, 190628, Leonli

    ReDim MTRString_TMPS(TheExec.sites.Existing.Count) 'Reset String Array
    ReDim MTRString_ROT(TheExec.sites.Existing.Count) 'Reset String Array
    ReDim MTRString_ROV(TheExec.sites.Existing.Count) 'Reset String Array

    Dim MTRTmpWave As New DSPWave
    Dim BeforeMTRTmpWave As New DSPWave
    Dim MTRDSPWave_ROT As New DSPWave
    Dim MTRDSPWave_ROV As New DSPWave
    Dim i As Integer        'Add, 190628, Leon li
    Dim CmdCheck As String
    Dim SweepCondition_Split() As String: SweepCondition_Split = Split(VddSenseFreq, "+")
    
    For i = 0 To UBound(SweepCondition_Split)
        If i = 0 Then
            'MTRDSPWave_ROT = GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & SensorName & "-sensor-ROT")
            MTRDSPWave_ROT = GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & SensorName & "--ROT")   'remove "sensor" for C-chop 20200107
            MTRDSPWave_ROV = GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & SensorName & "--ROV")   'remove "sensor" for C-chop 20200107
        Else
            For Each site In TheExec.sites
                MTRDSPWave_ROT = MTRDSPWave_ROT.Concatenate(GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & SensorName & "--ROT")) 'remove "sensor" for C-chop 20200107
                MTRDSPWave_ROV = MTRDSPWave_ROV.Concatenate(GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & SensorName & "--ROV")) 'remove "sensor" for C-chop 20200107
            Next site
        End If
    Next i
    For Each site In TheExec.sites
        MTRDSPWave_ROT = MTRDSPWave_ROT.Divide(1000000000#)
        MTRDSPWave_ROV = MTRDSPWave_ROV.Divide(1000000000#)
    Next site
    
    
    MTRTmpWave = GetStoredCaptureData(VddSenseHeat)
    BeforeMTRTmpWave = GetStoredCaptureData(VddSensePreheat)

'''    OffsetRecord = GetStoredMeasurement(VddVoltageLevelCNT)
    VoltageStepNumber = UBound(SweepCondition_Split) + 1
    For Each site In TheExec.sites.Active
        For i = 0 To VoltageStepNumber - 1
            MTRString_ROT(site) = MTRString_ROT(site) + CStr(FormatNumber(MTRDSPWave_ROT.Element(i), 8))
            MTRString_ROV(site) = MTRString_ROV(site) + CStr(FormatNumber(MTRDSPWave_ROV.Element(i), 8))
            MTRString_ROT(site) = MTRString_ROT(site) + " "
            MTRString_ROV(site) = MTRString_ROV(site) + " "
        Next i
        
        MTRString_ROT(site) = CmdROT + " " + SensorName + " " + MTRString_ROT(site)
        MTRString_ROV(site) = CmdROV + " " + SensorName + " " + MTRString_ROV(site)
        MTRString_TMPS(site) = CmdTMP + " " + SensorName + " " + CStr(FormatNumber((BeforeMTRTmpWave.Element(0) / 8), 7)) + " " + CStr(FormatNumber((MTRTmpWave.Element(0) / 8), 7))         ' Modify, 190628, Leon Li
    Next site
    


    SendCmd MTRString_ROT, CmdROT_Timeout
    SendCmd MTRString_ROV, CmdROV_Timeout
    SendCmd MTRString_TMPS, CmdTMP_Timeout     ' Modify, 190628, Leon Li
    
    CmdListStatus = 0 'Reset Command Result Status
    
    CmdCheck = CmdFinal + " " + SensorName + " -ts 10 40"                    ''''////remove for FT2 60C testing by 20190715 Leslie commend
    
    Set CmdListStatus = SendCmd(CmdCheck, CmpFinal_Timeout, False)
    TheExec.Flow.TestLimit CmdListStatus, 1, 1       ', , , , , , TestName
    
    
    '============Metrology UART Parser================================
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "****** RTOS MTR to EFUSE Hex2Decimal *******"
    
            Dim MTRFieldCount As Integer
            Dim rtos_mtr_fuse_name() As String     'Modify, Leon Li, 20190628
            Dim mm As Integer
            Dim rtos_mtr_fuse_value() As SiteLong     'Modify, Leon Li, 20190628
            Dim MTRFuseName As String

            MTRFieldCount = 15
            ReDim rtos_mtr_fuse_name(MTRFieldCount)
            For mm = 0 To MTRFieldCount

                  rtos_mtr_fuse_name(mm) = LCase("mtr_" & SensorName & "_c" & Trim(Str(mm)) & "=")
                  If mm = MTRFieldCount Then rtos_mtr_fuse_name(mm) = LCase("mtr_" & SensorName & "_ss=")
                  
            Next mm

            Dim strlen As Long
            Dim fuse_name_len As Long
            Dim fuse_code_idx As Long
            Dim fuse_code_value_hex() As New SiteVariant
            Dim fuse_code_value() As New SiteDouble
            Dim fuse_start_lo As Long
            Dim fuse_name_in_cate() As String
            
            ReDim fuse_code_value_hex(MTRFieldCount)
            ReDim fuse_code_value(MTRFieldCount)
            ReDim fuse_name_in_cate(MTRFieldCount)
            
            For Each site In TheExec.sites.Selected
              If CmdListStatus(site) = 1 Then
                    strlen = Len(GlobalMergeAry(site))
                    For mm = 0 To MTRFieldCount
                        If InStr(1, LCase(GlobalMergeAry(site)), rtos_mtr_fuse_name(mm)) <> 0 Then
                           
                                fuse_name_len = Len(rtos_mtr_fuse_name(mm))
                                fuse_code_idx = InStr(1, LCase(GlobalMergeAry(site)), rtos_mtr_fuse_name(mm))
                                fuse_start_lo = fuse_code_idx + fuse_name_len
                                fuse_name_in_cate(mm) = Replace(rtos_mtr_fuse_name(mm), "=", "")
                                fuse_code_value_hex(mm) = Mid(LCase(GlobalMergeAry(site)), fuse_start_lo, 10)
                                fuse_code_value(mm) = auto_HexStr2Value(fuse_code_value_hex(mm))     ' hex to dec
                                TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(site) + ") " + " MTR Hex EFUSE from UART                       " + Replace(rtos_mtr_fuse_name(mm), "=", " = ") + fuse_code_value_hex(mm)
                        
                        Else
                        
                                If InStr(1, LCase(GlobalMergeAry(site)), rtos_mtr_fuse_name(MTRFieldCount)) <> 0 Then         '''//// search "_ss" ////'''
                                        fuse_name_len = Len(rtos_mtr_fuse_name(MTRFieldCount))
                                        fuse_code_idx = InStr(1, LCase(GlobalMergeAry(site)), rtos_mtr_fuse_name(MTRFieldCount))
                                        fuse_start_lo = fuse_code_idx + fuse_name_len
                                        fuse_name_in_cate(MTRFieldCount) = Replace(rtos_mtr_fuse_name(MTRFieldCount), "=", "")
                                        fuse_code_value_hex(MTRFieldCount) = Mid(LCase(GlobalMergeAry(site)), fuse_start_lo, 10)
                                        fuse_code_value(MTRFieldCount) = auto_HexStr2Value(fuse_code_value_hex(MTRFieldCount))    ' hex to dec
                                        TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(site) + ") " + " MTR Hex EFUSE from UART                       " + Replace(rtos_mtr_fuse_name(MTRFieldCount), "=", " = ") + fuse_code_value_hex(MTRFieldCount)
                                End If
                            
                                If InStr(1, UCase(GlobalMergeAry(site)), UCase("ERROR")) <> 0 Then
                                        fuse_name_in_cate(mm) = "mtr_" & SensorName & "_c" & Trim(Str(mm))
                                        If mm = MTRFieldCount Then fuse_name_in_cate(mm) = "mtr_" & SensorName & "_ss"
                                        fuse_code_value(mm) = 0
                                        TheExec.sites(site).FlagState("F_Rtos_Metrology") = logicTrue
                                        GoTo NextForLoop:
                                End If
                                GoTo ExitForLoop:
                            
                        End If
NextForLoop:
                    Next mm
               Else
                '''//// For RTOS MTR fuse, Once MTRSNS fail, that will assign value into "0" ////'''
                    For mm = 0 To MTRFieldCount
                        fuse_name_in_cate(mm) = "mtr_" & SensorName & "_c" & Trim(Str(mm))
                        If mm = MTRFieldCount Then fuse_name_in_cate(mm) = "mtr_" & SensorName & "_ss"
                        fuse_code_value(mm) = 0
                    Next mm
                '''//// For RTOS MTR fuse, Once MTRSNS fail, that will assign value into "0" ////'''
               End If
ExitForLoop:
               TheExec.Datalog.WriteComment ""
            Next site
            
            For mm = 0 To MTRFieldCount
                Call AddStoredData(fuse_name_in_cate(mm), fuse_code_value(mm))
            Next mm
    '============Metrology UART Parser End===========================


    Dim instanceName As String
    instanceName = TheExec.DataManager.instanceName
    
    RTOS_UART_Print instanceName, CmdListStatus


Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function RTOS_Shmoo_Reboot(argc As Long, argv() As String) As Long


Dim LowPins As String
Dim HighPins As String
Dim HighToLow As String
Dim LowToHigh As String

Dim MyR As Variant
Dim TResult As New SiteLong
Dim dspData As New PinListData

Dim site As Variant
Dim LastPointFailed As Boolean

On Error GoTo errHandler
    LastPointFailed = False
    TheExec.Datalog.WriteComment "VDD_PCPU0:" & TheHdw.DCVS.Pins("VDD_PCPU0").Voltage.Value & "V"
    For Each site In TheExec.sites.Selected
        If Not (LastPointFailed) Then
            If (TheExec.DevChar.Results(argv(0)).Shmoo.CurrentPoint.ExecutionResult = tlDevCharResult_Fail) Then
                LastPointFailed = True
            End If
        End If
    Next site
    If (LastPointFailed) Then
        'RTOS_Boot_Up_fail_Power_Up 'Do power up again
        g_RTOS_FirstSetp = True
        If TheExec.EnableWord("RTOSRamp") = True Then
           RTOS_Boot False, , , True, True
        Else
           RTOS_Boot False, , , True, False
        End If
        
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function RTOS_Prepoint_check(argc As Long, argv() As String)
    Stop
End Function

 Public Function RTOS_RunScenario_ORI(Optional testName As String, Optional Cmd1 As String, Optional Cmd2 As String, Optional Cmd3 As String, Optional Cmd4 As String, _
            Optional Cmd5 As String, Optional Cmd1TimeOut As Double = 0#, Optional Cmd2TimeOut As Double = 0#, Optional Cmd3TimeOut As Double = 0#, _
            Optional Cmd4TimeOut As Double = 0#, Optional Cmd5TimeOut As Double = 0#, Optional SELSRAM_DSSC As String)
 
 On Error GoTo errHandler
    
     ReDim GlobalMergeAry(TheExec.sites.Existing.Count) 'for txt data collection
    
 '    Dim Cmd1Status As New SiteLong
 '    Dim Cmd2Status As New SiteLong
 '    Dim Cmd3Status As New SiteLong
 '    Dim Cmd4Status As New SiteLong
 '    Dim Cmd5Status As New SiteLong
    
     Dim CmdList As String
     Dim CmdListStatus As New SiteLong
    
     Dim CZSetupName As String
     Dim powerPin As String
     Dim SupplyVoltage As Long
     Dim TRef_Before_Pat As Double               '<- Code timing
     Dim TExec_Before_Pat As Double              '<- Execution time
     Shmoo_Pattern = testName
     Dim LogTimes As Boolean
   
 '    LogTimes = True
 '    If (LogTimes = True) Then
 '        TRef_Before_Pat = TheExec.Timer(0)
 '    End If
    
     Dim instanceName As String: instanceName = TheExec.DataManager.instanceName
     testName = instanceName
    
     If Cmd1 <> "" Then CmdList = Cmd1
     If Cmd2 <> "" Then CmdList = CmdList + ";" + Cmd2
     If Cmd3 <> "" Then CmdList = CmdList + ";" + Cmd3
     If Cmd4 <> "" Then CmdList = CmdList + ";" + Cmd4
     If Cmd5 <> "" Then CmdList = CmdList + ";" + Cmd5
    
     Dim Check_TestName As Long: Check_TestName = 0
     Dim Key As Variant: Key = Array("CHAR", "HBV")
     Dim Key_Count As Long
     For Key_Count = 0 To UBound(Key)
         Check_TestName = InStr(1, testName, Key(Key_Count)) 'Search the keyword from position 1
         If Check_TestName > 0 Then Check_TestName = Check_TestName + 1
     Next Key_Count
    
     If Check_TestName = 0 Then TheHdw.PinLevels.ApplyPower
        
     If Not (TheExec.DevChar.Setups.IsRunning) Then
         TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain '' use safe voltage to Selsram
     End If
    
     Shmoo_Save_core_power_per_site_for_Vbump ' store voltage into global variable
    
     'Select Sram Start
     Dim uniquesBit As Boolean, site As Variant
     uniquesBit = IsArray(SELSRAM_DSSC)
     Dim BinCutSelAry() As String
     BinCutSelAry = Split(SELSRAM_DSSC, ",")
    
     'If UBound(BinCutSelAry) = 5 And SELSRAM_DSSC <> "" Then
     If UBound(BinCutSelAry) = 4 And SELSRAM_DSSC <> "" Then 'for Tonga
         For Each site In TheExec.sites.Selected
             BinCutSelAry(site) = Decide_Switching_Bit_RTOS(BinCutSelAry(site), g_ApplyLevelTimingValt, "RTOS")
         Next site
         SendCmd BinCutSelAry, 0.1
     Else
         'SendCmd Decide_Switching_Bit_RTOS("SSSSS", g_ApplyLevelTimingValt, "RTOS"), 0.1
         SendCmd Decide_Switching_Bit_RTOS("SSSS", g_ApplyLevelTimingValt, "RTOS"), 0.1 ' for Tonga
     End If
    
     '//Change to Valt mode
     TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt '' only change the corepower to alt
     TheHdw.Wait 0.005
     
     '//Scenario Run Conditions
     CmdListStatus = 0
    
     If CmdList <> "" Then Set CmdListStatus = SendCmd(CmdList, Cmd1TimeOut, False)

     TheExec.Flow.TestLimit CmdListStatus, 1, 1 ', , , , , , TestName
     DebugPrintFunc ""
    
     RTOS_UART_Print instanceName, CmdListStatus
    
    
 '    For Each site In TheExec.sites.Active
 '        If (LogTimes = True) Then
 '            TExec_Before_Pat = TheExec.Timer(TRef_Before_Pat)
 '            TheExec.DataLog.WriteComment "ElapsedTime Pat Site (" & site & ")" + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
 '        End If
 '    Next site

     Exit Function
errHandler:
     If AbortTest Then Exit Function Else Resume Next

 End Function

Public Function RTOS_RunScenario(Optional testName As String, Optional Cmd1 As String, Optional Cmd2 As String, Optional Cmd3 As String, Optional Cmd4 As String, _
            Optional Cmd5 As String, Optional Cmd1TimeOut As Double = 0#, Optional Cmd2TimeOut As Double = 0#, Optional Cmd3TimeOut As Double = 0#, _
            Optional Cmd4TimeOut As Double = 0#, Optional Cmd5TimeOut As Double = 0#, Optional SELSRAM_DSSC As String, Optional Interpose_PrePat As String, Optional Pmode As String, Optional ForceCMD As String, Optional RampStep As Integer, Optional SetupCMD_Time As Double, Optional Vbump As Boolean = True, _
            Optional BinCutPowerPin As String, Optional Flag_Enable_BinCut_Rail_Switch As Boolean = False) As Long
            
On Error GoTo errHandler
    
    ReDim GlobalMergeAry(TheExec.sites.Existing.Count) 'for txt data collection
      
    Dim CmdList As Variant 'String
    Dim CmdListStatus As New SiteLong
    Dim CZSetupName As String
    Dim powerPin As String
    Dim SupplyVoltage As Long
    Dim TRef_Before_Pat As Double               '<- Code timing
    Dim TExec_Before_Pat As Double              '<- Execution time
    Shmoo_Pattern = testName
    Dim LogTimes As Boolean
    Dim instanceName As String: instanceName = TheExec.DataManager.instanceName
    Dim DevChar_Setup As String
    Dim CMDTotalTT As Double
    Dim BinCutSelAry() As String
    Dim SelsrmCmdForBincut As Boolean
    Dim RTOS_SELSRM_STR As String
    Vbump = True
    testName = instanceName
    
    If Cmd1 <> "" Then CmdList = Cmd1
    If Cmd2 <> "" Then CmdList = CmdList + Cmd2
    If Cmd3 <> "" Then CmdList = CmdList + Cmd3
    If Cmd4 <> "" Then CmdList = CmdList + Cmd4
    If Cmd5 <> "" Then CmdList = CmdList + Cmd5
    
    CMDTotalTT = Cmd1TimeOut + Cmd2TimeOut + Cmd3TimeOut + Cmd4TimeOut + Cmd5TimeOut
    If ForceCMD <> "" Then
        Call Replace_Force_cmd(CmdList, ForceCMD, CMDTotalTT)
    End If
    CMDTotalTT = CMDTotalTT + SetupCMD_Time
    CMDTotalTT = 5
    If Vbump = True Then 'SelSarm Funtion
        If Vbump = True Then g_Vbump_function = True    'add 20190625 by Leslie
        
            TheExec.EnableWord("RTOSRamp") = True
        
            If TheExec.DevChar.Setups.IsRunning = False Then
               g_RTOS_FirstSetp = True
            Else
               DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
               If TheExec.DevChar.Results(DevChar_Setup).startTime Like "1/1/0001*" Or TheExec.DevChar.Results(DevChar_Setup).startTime Like "0001/1/1*" Or g_LastRTOSPoint = True Then Exit Function
               g_LastRTOSPoint = False
            End If
        
        If g_RTOS_FirstSetp = True Then
           g_RTOSRampStep = 9
           If RampStep <> 0 Then
              If RampStep Mod 2 = 0 Then
                 g_RTOSRampStep = RampStep + 1
              Else
                 g_RTOSRampStep = RampStep
              End If
           End If
           TheHdw.PinLevels.ApplyPower
           Shmoo_Save_core_power_per_site_for_Vbump ' store voltage into global variable
           TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain '' use safe voltage to Selsram
           If Pmode <> "" Then
              If Interpose_PrePat <> "" Then
                 TheExec.ErrorLogMessage "Already exist Interpose_PrePat,Please check "
              Else
                 Decide_Pmode_ForceVoltage Pmode, "CorePower", Interpose_PrePat
              End If
           End If
           g_dyanmicDSSCbits = ""
            If SELSRAM_DSSC <> "" Then
                g_dyanmicDSSCbits = SELSRAM_DSSC
'               SELSRAM_DSSC = Replace((Replace(UCase(SELSRAM_DSSC), "'", "")), "SELSRM", "")
'               If UCase(SELSRAM_DSSC) Like "SSSSSSSSSSS" Then
'                  g_dyanmicDSSCbits = "SSSSSSSSSSS"
'               Else
'                  g_dyanmicDSSCbits = dynamic_SELSRM_source_bits(SELSRAM_DSSC, "RTOS")
'               End If
'            Else
'               g_dyanmicDSSCbits = "SSSSSSSSSSS"
                Shmoo_Save_core_power_per_site_for_Vbump
                For Each site In TheExec.sites.Active
                    RTOS_SELSRM_STR = Decide_Switching_Bit_RTOS(SELSRAM_DSSC, g_ApplyLevelTimingValt, "RTOS")
                    Exit For
                Next site
                SendCmd RTOS_SELSRM_STR, 0.1
                TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt
            Else
               g_dyanmicDSSCbits = "SSSSS"
            End If
        End If
    
        If Interpose_PrePat <> "" Then
           If g_RTOS_FirstSetp = True Then
             Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
             Getforcecondition_VDD g_ForceCond_VDD, Interpose_PrePat
           End If
        End If
        
        Dim Shmoo_Apply_Pin As String, pin_count As Long
        Get_Shmoo_Set_Pin Shmoo_Apply_Pin, g_ForceCond_VDD, pin_count
    
    
        'Select Sram Start
        Dim uniquesBit As Boolean ', site As Variant
        Dim RTOS_Char_SelAry() As Variant 'String
        ReDim RTOS_Char_SelAry(TheExec.sites.Existing.Count)
        
        For Each site In TheExec.sites.Selected
            RTOS_Char_SelAry(site) = Decide_Switching_Bit_RTOS(g_dyanmicDSSCbits, g_ApplyLevelTimingValt, "RTOS", Shmoo_Apply_Pin, g_Globalpointval, g_ForceCond_VDD, g_CharInputString_Voltage_Dict)
        Next site
         
        If g_RTOSNwireChar = True Then
           If g_RTOS2DFirstPoint = True Then
              RTOS_Boot False, , , True, False
           Else
              RTOS_Boot False, , , True, True
              g_RTOS_FirstSetp = True
           End If
        End If
        SendCmd RTOS_Char_SelAry, 0.1
    
        
        '//Change to Valt mode
        If g_RTOS_FirstSetp = True Then
           If TheExec.EnableWord("RTOSRamp") = True Then
              RTOS_Voltage_Rampdown
           Else
              Shmoo_Restore_Power_per_site_Vbump Shmoo_Apply_Pin
           End If
        Else
           g_VDDForce = ""
           Shmoo_Restore_Power_per_site_Vbump Shmoo_Apply_Pin
        End If
        g_RTOS_FirstSetp = False
        
        TheHdw.Wait 0.005
        
        'Scenario Run Conditions
        CmdListStatus = 0
        TheExec.Datalog.DatalogSuspended = False
    Else 'For Function Test and Bincut '''Bypass Selsram Function
        TheHdw.PinLevels.ApplyPower
        CmdListStatus = 0
        
'''        thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
''''        TheHdw.PinLevels.ApplyPower
'''        CmdListStatus = 0
'''        If Flag_Enable_BinCut_Rail_Switch = True Then
'''            g_TResultForBinCut = False
'''            SelsrmCmdForBincut = False
'''            For Each site In TheExec.sites.Selected
'''                'If Not UCase(Selsram_expand_array_forRTOS(site)) = "NONE" Then
'''                    SelsrmCmdForBincut = True
'''                    'BinCutSelAry(site) = Decide_Switching_Bit_RTOS(CStr(Selsram_expand_array_forRTOS(site)), , "RTOS", , , , , Flag_Enable_BinCut_Rail_Switch)
'''                    BinCutSelAry(site) = Decide_Switching_Bit_RTOS(g_dyanmicDSSCbits, g_ApplyLevelTimingValt, "RTOS", Shmoo_Apply_Pin, g_Globalpointval, g_ForceCond_VDD, g_CharInputString_Voltage_Dict)
'''                'End If
'''            Next site
'''            If SelsrmCmdForBincut = True Then SendCmd BinCutSelAry, 0.1
'''            thehdw.DCVS.Pins(BinCutPowerPin).Voltage.output = tlDCVSVoltageAlt
'''        ElseIf SELSRAM_DSSC <> "" Then
'''            Shmoo_Save_core_power_per_site_for_Vbump
'''            For Each site In TheExec.sites.Active
'''                RTOS_SELSRM_STR = Decide_Switching_Bit_RTOS(SELSRAM_DSSC, g_ApplyLevelTimingValt, "RTOS")
'''                Exit For
'''            Next site
'''            SendCmd RTOS_SELSRM_STR, 0.1
'''            thehdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt
'''        End If
    End If
    

    ' add this condition for BinCut, Flag_Enable_Rail_Switch=True, ==>Valt
    ' BinCut guys will decide on status of Flag_Enable_Rail_Switch
    If Flag_Enable_BinCut_Rail_Switch = True Then TheHdw.DCVS.Pins(BinCutPowerPin).Voltage.output = tlDCVSVoltageAlt
    
    g_TResultForBinCut = False ' initial status

    If CmdList <> "" Then Set CmdListStatus = SendCmd(CmdList, CMDTotalTT, False)

    If BinCutPowerPin = "" Then
        TheExec.Flow.TestLimit CmdListStatus, 1, 1       ', , , , , , TestName
    Else
        'do nothing  BinCut
    End If
    
    RTOS_UART_Print instanceName, CmdListStatus
    g_RTOS2DFirstPoint = False
    g_RTOSNwireChar = False
    If Vbump = True Then g_Vbump_function = False   'add 20190715 by Leslie
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function
Public Function RTOS_UART_Print(instanceName As String, Optional TResult As SiteLong)

On Error GoTo errHandler

    Dim site As Variant
    Dim PowerVolt As Double
    Dim powerPin As String
    Dim FName As String
    Dim OutputFilePath As String
    Dim day_code As String
    Dim SResult As String
    Dim CZSetupName As String
            
    Dim ByteCount As Long
    Dim asciiChar() As String
    Dim numericalVal() As Long
    
    Dim i As Long
    Dim j As Long
    
    Dim iPos As Long
    Dim FF_Count As Long
    Dim CR_Count As Long
            
'    Dim wb2 As Workbook: Set wb2 = Application.ActiveWorkbook
'    Dim ws_rtos As Worksheet: Set ws_rtos = wb2.Sheets("SpiromCodeFile")
'    Dim SPI_Versionint As String
'    Dim SPI_Version As String
'    SPI_Versionint = ws_rtos.Cells(1, 2).Value
    
'    If SPI_Versionint <> "" Then
'        SPI_Version = Replace(SPI_Versionint, ".\PATTERN\RTOS\SPIROM\", "")
'    Else
'        SPI_Version = ""
'    End If
    If TheExec.EnableWord("UARTOutPut") = True Then
        day_code = CStr(Year(Now)) & Right("0" & CStr(Month(Now)), 2) & Right("0" & CStr(day(Now)), 2)
        day_code = day_code & Right("0" & CStr(Hour(Now)), 2) & Right("0" & CStr(Minute(Now)), 2)
        
        For Each site In TheExec.sites
            If TResult(site) = 1 Then
                SResult = "Pass"
            Else
                SResult = "Fail"
            End If
            If TheExec.DevChar.Setups.IsRunning = True Then
                CZSetupName = TheExec.DevChar.Setups.ActiveSetupName
                powerPin = TheExec.DevChar.Setups(CZSetupName).Shmoo.Axes.Item(tlDevCharShmooAxis_X).ApplyTo.Pins
                'PowerVolt = Format(THEHDW.DCVS.Pins(powerPin).Voltage.Main.Value * 1000, "0")
                PowerVolt = Format(TheHdw.DCVS.Pins(powerPin).Voltage.Alt.Value * 1000, "0")
                OutputFilePath = ".\UART_Output\" & "Shmooing_Site" & site & "_" & "X_" & XCoord(site) & "_" & "Y_" & YCoord(site) & _
                                 "_" & instanceName & "_UARToutput_" & day_code & "_" & powerPin & "_" & CStr(PowerVolt) & "mV" & "_" & SResult & ".txt"
            Else
                'PowerVolt = Format(TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Main.Value * 1000, "0")
                OutputFilePath = ".\UART_Output\" & "Site" & site & "_" & "X_" & XCoord(site) & "_" & "Y_" & _
                                 YCoord(site) & "_" & instanceName & "_UARToutput_" & day_code & "_" & SResult & ".txt"
            End If
            
            Dim strlen As Long
            strlen = Len(GlobalMergeAry(site))
            
            FName = OutputFilePath
            Open FName For Append As #4
                Print #4, instanceName
                For i = 1 To strlen
                Dim TempStr As String
                TempStr = Mid(GlobalMergeAry(site), i, 1)
'                    Print #4, Mid(GlobalMergeAry(site), i, 1);
                    If Not (Asc(TempStr) = 255) Then
                        If Asc(TempStr) = 10 Or Asc(TempStr) = 13 Then  ''
                            Print #4, vbCrLf
                        Else
                            If Asc(TempStr) = 62 Then
                                Print #4, TempStr & vbCrLf;
    '                            Print #4, vbCrLf
                            Else
                                Print #4, TempStr;
                            End If
                        End If
                    End If
                Next i
            Close #4
            
        Next site
    End If
    ReDim GlobalMergeAry(TheExec.sites.Existing.Count)
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function SendCmdOrg(CmdStr As String, Optional ExtendedWait As Double = 0#, Optional CheckPassFail As Boolean = False, Optional LogUARTOutput As Boolean = True, Optional OutputFile As String) As SiteLong

    Dim CharArray() As String
    Dim dataArray() As Long
    Dim StringLength As Long
    Dim StringLengthPlus1 As Long
    Dim i As Long
    Dim dspData As New PinListData
    
'    Dim Site As Variant
    Dim DUTResponse As New DSPWave
    Dim TResult As New SiteLong
'    Dim ByteCount As Long
'
'    Dim asciiChar() As String
'    Dim numericalVal() As Long
'
'    Dim DateCodePath As String
'    Dim OutputFilePath As String
'    Dim temp_inst_name As String
'    temp_inst_name = TheExec.DataManager.InstanceName
'
'    Dim wb2 As Workbook: Set wb2 = Application.ActiveWorkbook
'    Dim ws_rtos As Worksheet: Set ws_rtos = wb2.Sheets("SpiromCodeFile")
'    Dim SPI_Versionint As String
'    Dim SPI_Version As String
'    SPI_Versionint = ws_rtos.Cells(1, 2).Value
'    If SPI_Versionint <> "" Then
'       SPI_Version = Replace(SPI_Versionint, ".\PATTERN\RTOS\SPIROM\", "")
'    Else
'       SPI_Version = ""
'    End If
'
'    Dim PowerVoltCPU, PowerVoltAON, PowerVoltWARM, PowerVoltDCS, PowerVoltSOC As Double
'    PowerVoltCPU = Format(thehdw.DCVS.Pins("VDD_CPU").Voltage.Main.Value * 1000, "0")
'    PowerVoltAON = Format(thehdw.DCVS.Pins("VDD_AON").Voltage.Main.Value * 1000, "0")
'    PowerVoltWARM = Format(thehdw.DCVS.Pins("VDD_WARM").Voltage.Main.Value * 1000, "0")
'    PowerVoltDCS = Format(thehdw.DCVS.Pins("VDD_DCS").Voltage.Main.Value * 1000, "0")
'    PowerVoltSOC = Format(thehdw.DCVS.Pins("VDD_SOC").Voltage.Main.Value * 1000, "0")
    
    StringLength = Len(CmdStr)
    StringLengthPlus1 = StringLength + 1
    
    ReDim CharArray(StringLength - 1)
    ReDim dataArray(StringLength)
'    ReDim dataArray(StringLengthPlus1)
    
    TheHdw.Protocol.ports("UART_RX").Enabled = True
    TheHdw.Protocol.ports("UART_TX").Enabled = True
   
    For i = 1 To StringLength
        CharArray(i - 1) = Mid(CmdStr, i, 1)
        dataArray(i - 1) = Asc(CharArray(i - 1))
    Next i
    dataArray(StringLength) = 13 ' carriage return
   ' dataArray(StringLengthPlus1) = 10 ' carriage return
TheHdw.Protocol.ports("UART_TX").NWire.MaxHoldUntilTimeout.Value = 0.007

'    Set dspData = thehdw.Protocol.ports("UART_TX").NWire.CMEM.dspWave
'
'    If ExtendedWait > 0.001 Then
'        thehdw.Protocol.ports("UART_TX").Modules("UART_read_response_extended").start
'    Else
'        thehdw.Protocol.ports("UART_TX").Modules("UART_read_response").start
'    End If
'    thehdw.Wait 0.02
     TheHdw.Protocol.ports("UART_TX").Modules("UART_read_response_extended").start '***
    
    For i = 0 To StringLength
    
'    For i = 0 To StringLengthPlus1 'StringLength
        With TheHdw.Protocol.ports("UART_RX").NWire.Frames("UART_Snd")
            .fields("Data_in").Value = dataArray(i)
            .Execute
        End With
        TheHdw.Protocol.ports("UART_RX").IdleWait
    Next i
    
    Set dspData = TheHdw.Protocol.ports("UART_TX").NWire.CMEM.DSPWave '***
'    thehdw.Protocol.ports("UART_TX").Modules("UART_read_response_extended").start
'
'    Else
'        thehdw.Protocol.ports("UART_TX").Modules("UART_read_response").start
'    End If
'
    If ExtendedWait > 0.001 Then
        TheHdw.Wait ExtendedWait
    Else
        TheHdw.Wait 0.02
    End If
    TheHdw.Protocol.ports("UART_TX").Enabled = False
    TheHdw.Protocol.ports("UART_RX").Enabled = False
    
'    OutputFilePath = ".\UART_Output\" & temp_inst_name & "_UARToutput_" & "CPU" & "_" & PowerVoltCPU & "_" & "AON" & "_" & PowerVoltAON & "_" & "WARM" & "_" & PowerVoltWARM & "_" & "DCS" & "_" & PowerVoltDCS & "_" & "SOC" & "_" & PowerVoltSOC & "_" & SPI_Version & ".txt"
    
'    If (CheckPassFail) Then
        If LogUARTOutput Then
            DUTResponse = dspData.Copy
            'TheExec.AddOutput "Capture count : " + Str(DUTResponse.SampleSize)
            'Compile_Error LogDUTResponse DUTResponse, OutputFile, CmdStr
'            LogDUTResponse DUTResponse, OutputFilePath, CmdStr
        End If
        
'        Dim Site As Variant
        

                'Compile_Error rundsp.ProcessDUTResponse dspData, TResult

        'Compile_Error Set SendCmd = TResult
'        TheExec.Flow.TestLimit TResult, 1, 1, , , , , , CmdStr, , , , , , , tlForceNone
'
'    Else
'        If LogUARTOutput Then
'            DUTResponse = dspData.Copy
'            LogDUTResponse DUTResponse, OutputFile, CmdStr
'        End If
'    End If
       

End Function



Public Function SendCmd(CmdStr As Variant, Optional ExtendedWait As Double = 0#, Optional CheckPassFail As Boolean = False) As SiteLong
''ZHHUANGF
On Error GoTo errHandler

    Dim CharArray() As String
    Dim dataArray() As Long
    
    Dim PerSiteDataArray() As New SiteLong
    
    Dim StringLength As Long
    Dim StringLengthPlus1 As Long
    Dim i As Long
    Dim dspData As New PinListData
    
    Dim DUTResponse As New DSPWave
    Dim TResult As New SiteLong
    
    Dim LongestCmd As Long
    Dim UniqueCmdPerSite As Boolean
    Dim site As Variant

    UniqueCmdPerSite = IsArray(CmdStr)
    LongestCmd = 0
    
    If UniqueCmdPerSite Then
        For Each site In TheExec.sites.Selected
            If LongestCmd < Len(CmdStr(site)) Then LongestCmd = Len(CmdStr(site))
        Next site
        StringLength = LongestCmd
    Else
        StringLength = Len(CmdStr)
    End If
    
    StringLengthPlus1 = StringLength + 1
    
    ReDim CharArray(StringLength - 1)
    ReDim dataArray(StringLength)
    ReDim PerSiteDataArray(StringLength)
    
'    ReDim dataArray(StringLengthPlus1)
    
    TheHdw.Protocol.ports("UART_RX").Enabled = True
    TheHdw.Protocol.ports("UART_TX").Enabled = True
    
    
    If UniqueCmdPerSite Then
        For Each site In TheExec.sites.Selected
            For i = 1 To StringLength
                If i <= Len(CmdStr(site)) Then
                    CharArray(i - 1) = Mid(CmdStr(site), i, 1)
                    PerSiteDataArray(i - 1) = Asc(CharArray(i - 1))
                Else
                    PerSiteDataArray(i - 1) = Asc(" ")
                End If
            Next i
            PerSiteDataArray(StringLength) = 13 ' carriage return
        Next site
    Else
        For i = 1 To StringLength
            CharArray(i - 1) = Mid(CmdStr, i, 1)
            dataArray(i - 1) = Asc(CharArray(i - 1))
        Next i
        dataArray(StringLength) = 13 ' carriage return
        ' dataArray(StringLengthPlus1) = 10 ' carriage return
    End If
    
    TheHdw.Protocol.ports("UART_TX").NWire.MaxHoldUntilTimeout.Value = 0.11
    
    TheHdw.Protocol.ports("UART_TX").NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
    
    TheHdw.Protocol.ports("UART_TX").Modules("UART_read_response_extended").start '***
    
    If UniqueCmdPerSite Then
        For i = 0 To StringLength
            With TheHdw.Protocol.ports("UART_RX").NWire.Frames("UART_Snd")
                .fields("Data_in").Value = PerSiteDataArray(i)
                .Execute
            End With
            TheHdw.Protocol.ports("UART_RX").IdleWait
        Next i
    Else
        For i = 0 To StringLength
            With TheHdw.Protocol.ports("UART_RX").NWire.Frames("UART_Snd")
                .fields("Data_in").Value = dataArray(i)
                .Execute
            End With
            TheHdw.Protocol.ports("UART_RX").IdleWait
        Next i
    End If
    
    Set dspData = TheHdw.Protocol.ports("UART_TX").NWire.CMEM.DSPWave '***

    If ExtendedWait > 0.001 Then
        TheHdw.Wait ExtendedWait
    Else
        TheHdw.Wait 0.02
    End If
    TheHdw.Protocol.ports("UART_TX").Enabled = False
    TheHdw.Protocol.ports("UART_RX").Enabled = False
'----------------------------------------add to avoid 2 interations-----------------------
    Dim Prompt_Idx As New SiteLong

             If TheExec.DataManager.instanceName Like "*S094*" Or TheExec.DataManager.instanceName Like "*S095*" Or TheExec.DataManager.instanceName Like "*S096*" Then
                Prompt_Idx = 0
             Else
                Prompt_Idx = 1
             End If



    
    
    
    
    
    '9th, Oct 2019
    'After Discussing with customer,TER-Fred add, for TP checking, bypass dsp process error
    '==================================================================Start
    If TheExec.TesterMode = testModeOffline Then
            Dim PassArrayOffline(8) As Long
            Dim FailArrayOffline(8) As Long
            Dim dspdata_checking As New DSPWave
            
            Dim VarSite As Variant
            Dim RandomVal As Double
        


            For Each VarSite In TheExec.sites.Selected
                RandomVal = Rnd      '0<=RandomVal<1
                If RandomVal < 0.8 Or EnableWord_Golden_Default = True Then
                    PassArrayOffline(0) = 80 'P
                    PassArrayOffline(1) = 65 'A
                    PassArrayOffline(2) = 83 'S
                    PassArrayOffline(3) = 83 'S
                    PassArrayOffline(4) = CLng(Asc(" ")) 'space
                    PassArrayOffline(5) = 65 'A
                    PassArrayOffline(6) = 84 'T
                    PassArrayOffline(7) = 69 'E
                    PassArrayOffline(8) = 62 '>
                
                    dspdata_checking.Data = PassArrayOffline
                Else
                    FailArrayOffline(0) = 70 'F
                    FailArrayOffline(1) = 65 'A
                    FailArrayOffline(2) = 73 'I
                    FailArrayOffline(3) = 76 'L
                    FailArrayOffline(4) = CLng(Asc(" ")) 'space
                    FailArrayOffline(5) = 65 'A
                    FailArrayOffline(6) = 84 'T
                    FailArrayOffline(7) = 69 'E
                    FailArrayOffline(8) = 62 '>
                
                    dspdata_checking.Data = FailArrayOffline
                End If
            Next VarSite

            Dim UseCmdLength_1 As Boolean
            rundsp.ProcessDUTResponse dspdata_checking, TResult, Prompt_Idx, StringLength, UseCmdLength_1
    
            For Each VarSite In TheExec.sites.Selected
                DUTResponse(VarSite) = dspdata_checking(VarSite).Copy
            Next VarSite
    Else

'below code are the original vbt
       Dim UseCmdLength As Boolean
      '  rundsp.ProcessDUTResponse dspData, TResult, Prompt_Idx
                rundsp.ProcessDUTResponse dspData, TResult, Prompt_Idx, StringLength, UseCmdLength
                'rundsp.ProcessDUTResponse dspData, TResult, Prompt_Idx, Len(CmdStr), UseCmdLength 'up line Len(CmdStr) change to "StringLength" to avoid type mismatch if "CmdStr" is array
    'rundsp.ProcessDUTResponse dspData, TResult
'---------------------------------------------------------------------------------------
    DUTResponse = dspData.Copy
    
    End If
    Call LogDUTResponse(DUTResponse, TResult)

    Set SendCmd = TResult
        
        '---------------------------------------------------------------------------------------
    ' add a global parameter "g_TResultForBinCut" for Bincut
    For Each site In TheExec.sites
        If TResult = 1 Then
            g_TResultForBinCut = True
        Else
            g_TResultForBinCut = False
        End If
    Next site
'---------------------------------------------------------------------------------------
    
    If CheckPassFail Then
        TheExec.Flow.TestLimit TResult, 1, 1, , , , , , CmdStr, , , , , , , tlForceNone
    End If
'
       
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function LogDUTResponse(DUTResponse As DSPWave, Optional TResult As SiteLong, Optional OutputToDebuglog As Boolean, Optional OutputToDatalog As Boolean)

    Dim site As Variant
    Dim ByteCount As Long

    Dim asciiChar() As String
    Dim numericalVal() As Long
    
    Dim i As Long
    Dim j As Long
    
    Dim iPos As Long
    Dim FF_Count As Long
    Dim CR_Count As Long
    Dim print_file_cmdStr As String


    OutputToDatalog = False: OutputToDebuglog = False
    OutputToDatalog = True
    If TheExec.Flow.EnableWord("UARTToDatalog") = True Then OutputToDatalog = True
    If TheExec.Flow.EnableWord("UARTToDebuglog") = True Then OutputToDebuglog = True
    
    ReDim Preserve GlobalMergeAry(TheExec.sites.Existing.Count)
    
    For Each site In TheExec.sites.Selected
        j = 0
        ByteCount = DUTResponse.SampleSize
        
        ReDim asciiChar(ByteCount - 1)
        ReDim numericalVal(ByteCount - 1)
        

        For i = 0 To ByteCount - 1
            asciiChar(i) = ""
            numericalVal(i) = DUTResponse.Element(i)
            If (numericalVal(i) <> 255) And (numericalVal(i) <> 10) Then
                asciiChar(j) = Chr(numericalVal(i))
                j = j + 1
            End If
        Next i
        
        GlobalMergeAry(site) = GlobalMergeAry(site) & Join(asciiChar(), "")
        
        If OutputToDatalog Then
            TheExec.Datalog.WriteComment "****************************************"
            TheExec.Datalog.WriteComment "Site : " + CStr(site)
            TheExec.Datalog.WriteComment "****************************************"
            
            WriteToDatalog asciiChar, j
    
        ElseIf OutputToDebuglog Then
'            CR_Count = 0
'            FF_Count = 0
'            For i = 0 To ByteCount - 1
''                numericalVal(i) = DUTResponse(site).Element(i)
'                numericalVal(i) = DUTResponse.Element(i)
'                If numericalVal(i) = 255 Then
'                FF_Count = FF_Count + 1
'                    numericalVal(i) = Asc("#")
'                End If
'
'                If numericalVal(i) = 0 Then numericalVal(i) = Asc("@")
'                If numericalVal(i) = 10 Then CR_Count = CR_Count + 1
'                If (numericalVal(i) <> 255) And (numericalVal(i) <> 10) Then
'                    asciiChar(j) = Chr(numericalVal(i))
'                    j = j + 1
'               TheExec.Datalog.WriteComment CStr(i) + " " + CStr(numericalVal(i)) + " " + Chr(numericalVal(i))
'                End If
'            Next i
''
'                TheExec.AddOutput "FF count : " + str(FF_Count)
'                TheExec.AddOutput "CR count : " + str(CR_Count)
                        
                TheExec.AddOutput "****************************************"
                TheExec.AddOutput "Site : " + CStr(site)
                TheExec.AddOutput "****************************************"
                WriteToOutputWindow asciiChar, j
        End If
        
    Next site

End Function

Public Sub SendCmdOnly(CmdStr As String)

    Dim CharArray() As String
    Dim StringLength As Long
    Dim i As Long
    

    Dim dataArray() As Integer
    
    StringLength = Len(CmdStr)
    ReDim CharArray(StringLength - 1)
    ReDim dataArray(StringLength)
    
   
    For i = 1 To StringLength
        CharArray(i - 1) = Mid(CmdStr, i, 1)
        dataArray(i - 1) = Asc(CharArray(i - 1))
    Next i
    dataArray(StringLength) = 13 ' carriage return

    For i = 0 To StringLength    ' leave newline char to read module
        With TheHdw.Protocol.ports("UART_RX").NWire.Frames("UART_Snd")
            .fields("Data_in").Value = dataArray(i)
            .Execute
        End With
'        thehdw.Protocol.ports("UART_PA").IdleWait
    Next i
         

End Sub


Public Sub WriteToOutputWindow(dataArray() As String, CharCount As Long)

Dim OutLine As String
Dim i As Long


    OutLine = ""
    For i = 0 To CharCount - 1  'ubound(dataarray)-1
'    If i = 157 Then
'    Stop
'    End If
'Debug.Print DataArray(i)
        If (Asc(dataArray(i)) = 10) Or (Asc(dataArray(i)) = 13) Then
            TheExec.AddOutput OutLine
            OutLine = ""
        Else
            OutLine = OutLine + dataArray(i)
'            If (Asc(dataArray(i)) = 62) Then
'                theexec.AddOutput OutLine
'                OutLine = ""
'                i = CharCount   'UBound(DataArray)
'            End If
            
        End If
    Next i
    
    If Len(OutLine) > 0 Then
        TheExec.AddOutput OutLine
    End If
        
End Sub

Public Sub WriteToDatalog(dataArray() As String, CharCount As Long)

Dim OutLine As String
Dim i As Long


    OutLine = ""
    For i = 0 To CharCount - 1  'ubound(dataarray)-1
'    If i = 157 Then
'    Stop
'    End If
'Debug.Print DataArray(i)
        If (Asc(dataArray(i)) = 10) Or (Asc(dataArray(i)) = 13) Then
            TheExec.Datalog.WriteComment OutLine
            OutLine = ""
        Else
            OutLine = OutLine + dataArray(i)
'            If (Asc(dataArray(i)) = 62) Then          '20200107 to avoid ">" then exit function,due to future project will output many ">"
'                TheExec.Datalog.WriteComment OutLine
'                OutLine = ""
'                i = CharCount   'UBound(DataArray)
'            End If
            
        End If
    Next i
    
    If Len(OutLine) > 0 Then
        TheExec.Datalog.WriteComment OutLine
    End If
        
End Sub

Public Function ReloadUARTModules() As Long

On Error GoTo errHandler
    
    TheHdw.Protocol.ports("UART_TX").ModuleFiles.UnloadAll
    TheHdw.Protocol.ports("UART_RX").ModuleFiles.UnloadAll
    
    TheHdw.Protocol.ports("UART_TX").Enabled = True
    TheHdw.Protocol.ports("UART_RX").Enabled = True
    
    TheHdw.Protocol.ports("UART_TX").NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
   
    With TheHdw.Protocol.ports("UART_TX")
        If (Not .ModuleFiles.Contains("VBT_UART_TX_module")) Then
            Call .ModuleFiles.Load("VBT_UART_TX_module")
        End If
    End With
        
    With TheHdw.Protocol.ports("UART_RX")
        If (Not .ModuleFiles.Contains("VBT_UART_RX_module")) Then
            Call .ModuleFiles.Load("VBT_UART_RX_module")
        End If
    End With
    
    TheHdw.Protocol.ports("UART_TX").Enabled = False
    TheHdw.Protocol.ports("UART_RX").Enabled = False
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RTOS_RunScenario_mod(Optional testName As String, Optional Cmd1 As String, Optional Cmd1TimeOut As Double = 0#, Optional SelsramBit As String) As SiteLong

On Error GoTo errHandler
    
    ReDim GlobalMergeAry(TheExec.sites.Existing.Count) 'for txt data collection
    
'    Dim Cmd1Status As New SiteLong
'    Dim Cmd2Status As New SiteLong
'    Dim Cmd3Status As New SiteLong
'    Dim Cmd4Status As New SiteLong
'    Dim Cmd5Status As New SiteLong
    
    Dim CmdList As String
    Dim CmdListStatus As New SiteLong
    
    Dim CZSetupName As String
    Dim powerPin As String
    Dim SupplyVoltage As Long
    Dim TRef_Before_Pat As Double               '<- Code timing
    Dim TExec_Before_Pat As Double              '<- Execution time
    Shmoo_Pattern = testName
    Dim LogTimes As Boolean
   
'    LogTimes = True
'    If (LogTimes = True) Then
'        TRef_Before_Pat = TheExec.Timer(0)
'    End If
    
    Dim instanceName As String: instanceName = TheExec.DataManager.instanceName
    ''TestName = instanceName
    
    If Cmd1 <> "" Then CmdList = Cmd1
    
    Dim Check_TestName As Long: Check_TestName = 0
        
    Shmoo_Save_core_power_per_site_for_Vbump ' store voltage into global variable
    
    'Select Sram Start
    Dim uniquesBit As Boolean, site As Variant
    uniquesBit = IsArray(SelsramBit)
    Dim BinCutSelAry() As String
    BinCutSelAry = Split(SelsramBit, ",")
    
    'If UBound(BinCutSelAry) = 5 And SelsramBit <> "" Then
    If UBound(BinCutSelAry) = 4 And SelsramBit <> "" Then 'for Tonga
        For Each site In TheExec.sites.Selected
            BinCutSelAry(site) = Decide_Switching_Bit_RTOS(BinCutSelAry(site), g_ApplyLevelTimingValt, "RTOS")
        Next site
        SendCmd BinCutSelAry, 0.1
    Else
        'SendCmd Decide_Switching_Bit_RTOS("SSSSS", g_ApplyLevelTimingValt, "RTOS"), 0.05
        SendCmd Decide_Switching_Bit_RTOS("SSSS", g_ApplyLevelTimingValt, "RTOS"), 0.05 'for Tonga
    End If
    
    '//Change to Valt mode
    TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt '' only change the corepower to alt
    TheHdw.Wait 0.005
    'Scenario Run Conditions

    CmdListStatus = 0
    
    If CmdList <> "" Then Set CmdListStatus = SendCmd(CmdList, Cmd1TimeOut, False)

    'theexec.Flow.TestLimit CmdListStatus, 0, 0       ', , , , , , TestName
    Set RTOS_RunScenario_mod = CmdListStatus
    
    RTOS_UART_Print instanceName, CmdListStatus
    
    
'    For Each site In TheExec.sites.Active
'        If (LogTimes = True) Then
'            TExec_Before_Pat = TheExec.Timer(TRef_Before_Pat)
'            TheExec.DataLog.WriteComment "ElapsedTime Pat Site (" & site & ")" + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
'        End If
'    Next site

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function Replace_Force_cmd(CmdList As Variant, Force_cmd As String, CmdWaitTime As Double) As Double

On Error GoTo errHandler

'    Dim PlaceGLCLocation() As String
    Dim tmpStr() As String
    Dim SC_array() As String
    Dim CoreMask_array() As String
    Dim GLC_Val As String
    Dim CoreMask_Val As String
    Dim Flavor_Val As String
    Dim i, j, k As Integer
    Dim w_idx As Integer
    Dim GLC_WT_idx As Integer
    Dim TmpVal As Long
    Dim TmpStr_1 As String
    Dim CmdLsitArray() As String
    Dim CmdWaitTimeArray() As String
    Dim GLC_Ratio As Double
    
    
    
    tmpStr = Split(Force_cmd, " ")
    For i = 0 To UBound(tmpStr)
        If UCase(tmpStr(i)) Like "GLP*" Then
            GLC_Val = CStr(Mid(tmpStr(i), 4))
        ElseIf UCase(tmpStr(i)) Like "*COREMASK*" Then
            CoreMask_Val = CStr(Mid(tmpStr(i), 9))
        ElseIf UCase(tmpStr(i)) Like "FLAVOR*" Then
            Flavor_Val = CStr(Mid(tmpStr(i), 7))
        End If
    Next i
    
    
    CmdLsitArray = Split(CmdList, ";")
'    CmdWaitTimeArray = Split(CmdWaitTime, ";")
    
    For i = 0 To UBound(tmpStr)
        If UCase(tmpStr(i)) Like "GLP*" Then
            w_idx = InStr(UCase(CmdList), "SC RUN")
            If InStr(Mid(CmdList, w_idx), ";") = 0 Then
                SC_array = Split(Mid(CmdList, w_idx), " ")
            Else
                SC_array = Split(Mid(CmdList, w_idx, InStr(Mid(CmdList, w_idx), ";") - 1), " ")
            End If
            If IsNumeric(SC_array(3)) = True Then ' append new GLC value
                GLC_Ratio = CDbl(Format((CDbl(GLC_Val) / CDbl(SC_array(3))), "0.000"))
                SC_array(3) = GLC_Val
            Else
                SC_array(2) = SC_array(2) & " " & GLC_Val ' use new GLC value from Force_CMD
            End If
            For j = 0 To UBound(CmdLsitArray)
                If UCase(CmdLsitArray(j)) Like "*SC RUN*" Then
                    CmdLsitArray(j) = Join(SC_array)
                    Exit For
                End If
            Next j
        ElseIf UCase(tmpStr(i)) Like "*COREMASK*" Then
            w_idx = InStr(UCase(CmdList), "-COREMASK")
            If InStr(Mid(CmdList, w_idx), ";") = 0 Then
                CoreMask_array = Split(Mid(CmdList, w_idx), " ")
            Else
                CoreMask_array = Split(Mid(CmdList, w_idx, InStr(Mid(CmdList, w_idx), ";") - 1), " ")
            End If
            CoreMask_array(1) = CoreMask_Val
            For j = 0 To UBound(CmdLsitArray)
                If UCase(CmdLsitArray(j)) Like "*COREMASK*" Then
                    CmdLsitArray(j) = Join(CoreMask_array)
                    Exit For
                End If
            Next j

        End If
    Next i
    CmdList = Join(CmdLsitArray, ";")
    
    CmdWaitTime = CmdWaitTime * GLC_Ratio

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function RTOS_Boot_and_CSW(BootUsingPattern As Boolean, Optional BootPattern As Pattern, Optional UseJTAG As Boolean, Optional shmooing As Boolean, Optional ramp As Boolean = False) As Long

On Error GoTo errHandler
    
    If Not (shmooing) Then
        ReDim GlobalMergeAry(TheExec.sites.Existing.Count)
    End If
    
    Dim Relay_Device As String
    Dim Relay_Spirom As String
    Dim dspData As New PinListData
    Dim LowPins As String
    Dim HighPins As String
    
    Dim LowToHigh As String
    Dim HighToLow As String

    Dim BootDSP As New DSPWave
    Dim i As Long, p As Long
    Dim TResult As New SiteLong
    Dim LogTimes As Boolean
    Dim TRef_Before_Pat As Double               '<- Code timing
    Dim TExec_Before_Pat As Double              '<- Execution time
    Dim instanceName As String
    Dim Volt_point As Integer
    instanceName = TheExec.DataManager.instanceName
    TNTEMP = instanceName
    '======================================================
    Relay_Device = "k36,k38" 'update for Tonga ''''' "k01,k03"
    Relay_Spirom = "k37,k39" ' "k02,k04"
    LowPins = "RTOS_Boot_Low"    'Pin group"
    HighPins = "RTOS_Boot_High"  'Pin group"
    HighToLow = ""
    LowToHigh = "RTOS_Boot_L2H"  'Pin group
    '======================================================
    
'    LogTimes = True
'    If (LogTimes = True) Then
'        TRef_Before_Pat = TheExec.Timer(0)
'    End If
 
    Call UpdateDLogColumns(40)
    
    If ramp = True Then
       RTOS_Voltage_RampUp
    Else
       TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If

   'For Volt_point = 0 To 2
            TheHdw.Utility.Pins(Relay_Spirom).State = tlUtilBitOff
            TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOff
            TheHdw.Wait 0.003
                
            TheHdw.DCVS.Pins("SPI_PWR").Connect                     ' re-cycle SPI-ROM power
            TheHdw.DCVS.Pins("SPI_PWR").Voltage.Main.Value = 0#
            TheHdw.Wait 0.02
            TheHdw.DCVS.Pins("SPI_PWR").Voltage.Main.Value = 1.8
            TheHdw.Wait 0.01
        
'            thehdw.Digital.Pins("xo0,jtag_tck").Disconnect
'            Start_Profile_AutoResolution "SPI_PWR", "I", 0, 0, "RTOS", 1

'''//Follow relay switch for Turks//'''
    TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
    TheHdw.Wait 0.05
    

            With TheHdw.Protocol.ports("UART_TX")
                .TimeOut.Enabled = True
                .TimeOut.Value = 2
                .Enabled = True
                .NWire.MaxHoldUntilTimeout.Value = 0.003    ''''////3msec*1500=4.5sec////''''
                '.NWire.MaxWaitUntilTimeout.Value = 0.005
                .NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
            End With
            Set dspData = TheHdw.Protocol.ports("UART_TX").NWire.CMEM.DSPWave
             
            TheHdw.Protocol.ports("UART_TX").Modules("UART_boot").start
            
            TheHdw.Wait 0.002
              
              
              
            If BootUsingPattern Then
                TheHdw.Patterns(BootPattern).Load
    
            If UseJTAG Then
    '''            TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
                TheHdw.Wait 0.005
                TheHdw.Patterns(BootPattern).start
                TheHdw.Digital.Patgen.HaltWait
            Else
                TheHdw.Digital.Patgen.Continue 0, cpuA
                TheHdw.Patterns(BootPattern).start
                TheHdw.Digital.Patgen.FlagWait cpuA, 0
                TheHdw.Wait 0.003
    '''            TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
                TheHdw.Wait 0.01
                TheHdw.Digital.Patgen.Continue 0, cpuA
                TheHdw.Digital.Patgen.HaltWait
            End If
        Else
            TheHdw.Digital.Pins(LowPins).InitState = chInitLo
            TheHdw.Digital.Pins(HighPins).InitState = chInitHi
            TheHdw.Wait 0.01
    '''        TheHdw.Digital.Pins("SPI1_MISO").InitState = chInitHi ''001 -> 101   '''//Cebu has
    '''        TheHdw.Wait 0.05
        
            TheHdw.Digital.Pins(LowToHigh).InitState = chInitHi
            TheHdw.Wait 0.05
            
    ''''////Follow relay switch for Cebu//''''
    '''        TheHdw.Utility.Pins(Relay_Device).State = tlUtilBitOn
    '''        TheHdw.Wait 0.05
    
    ''''////==== For JTAG Debug ====////''''
    '''    thehdw.Digital.Pins("jtag_tdi,jtag_tdo,jtag_sel,jtag_trstn,jtag_tms, jtag_tck").Disconnect '<+For JTAG mode?
    '''    thehdw.Wait 0.01
    
        End If
    
        '    Plot_Profile "SPI_PWR", "RTOS"
        
            TheHdw.Protocol.ports("UART_TX").IdleWait
            TheHdw.Protocol.ports("UART_TX").Enabled = False
        
        
            rundsp.CheckBootStatus dspData, TResult  'Check DSP wave status to determine TResult
                    
            BootDSP = dspData.Copy
            
            Call LogDUTResponse(BootDSP, TResult) 'Copy DSP wave into an output log
                        
            TheHdw.Protocol.ports("UART_TX").TimeOut.Value = 30#
            
             'Boot up Configeration
            TResult = SendCmd("slave up acc", 0.2)  ''''////Plain boot////''''
        
        
            If Not (shmooing) Then RTOS_UART_Print instanceName, TResult
            If Not (shmooing) Then TheExec.Flow.TestLimit TResult, 1, 1, , , , , , "Boot Status"
           

            TheHdw.Wait 0.01

           If LCase(instanceName) Like "*450mv*" Then
               TheHdw.DCVS.Pins("VDD_ECPU").Voltage.Value = 0.45
               TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Value = 0.45
               TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Value = 0.75
               force_val = 0.45
            ElseIf LCase(instanceName) Like "*550mv*" Then
               TheHdw.DCVS.Pins("VDD_ECPU").Voltage.Value = 0.55
               TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Value = 0.55
               TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Value = 0.75
               force_val = 0.55
            ElseIf LCase(instanceName) Like "*650mv*" Then
               TheHdw.DCVS.Pins("VDD_ECPU").Voltage.Value = 0.65
               TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Value = 0.65
               TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Value = 0.75
               force_val = 0.65
            ElseIf LCase(instanceName) Like "*750mv*" Then
               TheHdw.DCVS.Pins("VDD_ECPU").Voltage.Value = 0.75
               TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Value = 0.75
               TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Value = 0.75
               force_val = 0.75
            ElseIf LCase(instanceName) Like "*850mv*" Then
               TheHdw.DCVS.Pins("VDD_ECPU").Voltage.Value = 0.85
               TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Value = 0.85
               TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Value = 0.85
               force_val = 0.85
            ElseIf LCase(instanceName) Like "*950mv*" Then
               TheHdw.DCVS.Pins("VDD_ECPU").Voltage.Value = 0.95
               TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Value = 0.95
               TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Value = 0.95
               force_val = 0.95
            ElseIf LCase(instanceName) Like "*1050mv*" Then
               TheHdw.DCVS.Pins("VDD_ECPU").Voltage.Value = 1.05
               TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Value = 1.05
               TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Value = 1.05
               force_val = 1.05
            ElseIf LCase(instanceName) Like "*1150mv*" Then
               TheHdw.DCVS.Pins("VDD_ECPU").Voltage.Value = 1.15
               TheHdw.DCVS.Pins("VDD_PCPU").Voltage.Value = 1.15
               TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Value = 1.15
               force_val = 1.15
            End If

            PinTEMP = "VDD_PCPU,VDD_ECPU,VDD_CPU_SRAM"
            TheHdw.DCVS.Pins("VDD_PCPU").SetCurrentRanges 30, 30
            TheHdw.DCVS.Pins("VDD_ECPU").SetCurrentRanges 15, 15
            TheHdw.DCVS.Pins("VDD_CPU_SRAM").SetCurrentRanges 15, 15
            
            
          PinTEMP = "VDD_PCPU"
          SendCmd_CSW "pmgr mode 0xD222134", 0.05
          TNTEMP = TNTEMP + "_0xD222134"
          SendCmd_CSW "sc run 37", 0.25
          TNTEMP = instanceName

          SendCmd_CSW "pmgr mode 0xE222134", 0.05
          TNTEMP = TNTEMP + "_0xE222134"
          SendCmd_CSW "sc run 37", 0.25
          TNTEMP = instanceName
        
          SendCmd_CSW "pmgr mode 0xF222134", 0.05
          TNTEMP = TNTEMP + "_0xF222134"
          SendCmd_CSW "sc run 37", 0.25
          TNTEMP = instanceName
          
          
          TheHdw.Wait 0.01
        
        
          PinTEMP = "VDD_ECPU"
          SendCmd_CSW "pmgr mode 0x2D22134", 0.05
          TNTEMP = TNTEMP + "_0x2D22134"
          SendCmd_CSW "sc run 36", 0.25
          TNTEMP = instanceName

          SendCmd_CSW "pmgr mode 0x2E22134", 0.05
          TNTEMP = TNTEMP + "_0x2E22134"
          SendCmd_CSW "sc run 36", 0.25
          TNTEMP = instanceName
        
          SendCmd_CSW "pmgr mode 0x2F22134", 0.05
          TNTEMP = TNTEMP + "_0x2F22134"
          SendCmd_CSW "sc run 36", 0.25
          TNTEMP = instanceName
            

         If Not (shmooing) Then RTOS_UART_Print instanceName, TResult
               
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function SendCmd_CSW(CmdStr As Variant, Optional ExtendedWait As Double = 0#, Optional CheckPassFail As Boolean = False) As SiteLong

On Error GoTo errHandler
    Dim measureCurrent As New PinListData
    Dim HexmeasureCurrent As New PinListData
    Dim UvsmeasureCurrent As New PinListData
    Dim pintempary() As String
    Dim uvsary As String
    Dim hexary As String
    Dim index As Integer
    Dim i As Long
    
    pintempary = Split(PinTEMP, ",")
    
    For i = 0 To UBound(pintempary)
            hexary = pintempary(i)
            HexmeasureCurrent.AddPin (hexary)
    Next i

    Dim CharArray() As String
    Dim dataArray() As Long
    
    Dim PerSiteDataArray() As New SiteLong
    
    Dim StringLength As Long
    Dim StringLengthPlus1 As Long
    Dim dspData As New PinListData
    
    Dim DUTResponse As New DSPWave
    Dim TResult As New SiteLong
    
    Dim LongestCmd As Long
    Dim UniqueCmdPerSite As Boolean
    Dim site As Variant

    UniqueCmdPerSite = IsArray(CmdStr)
    LongestCmd = 0
    
    If UniqueCmdPerSite Then
        For Each site In TheExec.sites.Selected
            If LongestCmd < Len(CmdStr(site)) Then LongestCmd = Len(CmdStr(site))
        Next site
        StringLength = LongestCmd
    Else
        StringLength = Len(CmdStr)
    End If
    
    StringLengthPlus1 = StringLength + 1
    
    ReDim CharArray(StringLength - 1)
    ReDim dataArray(StringLength)
    ReDim PerSiteDataArray(StringLength)
    
'    ReDim dataArray(StringLengthPlus1)
    
    TheHdw.Protocol.ports("UART_RX").Enabled = True
    TheHdw.Protocol.ports("UART_TX").Enabled = True
    
    
    If UniqueCmdPerSite Then
        For Each site In TheExec.sites.Selected
            For i = 1 To StringLength
                If i <= Len(CmdStr(site)) Then
                    CharArray(i - 1) = Mid(CmdStr(site), i, 1)
                    PerSiteDataArray(i - 1) = Asc(CharArray(i - 1))
                Else
                    PerSiteDataArray(i - 1) = Asc(" ")
                End If
            Next i
            PerSiteDataArray(StringLength) = 13 ' carriage return
        Next site
    Else
        For i = 1 To StringLength
            CharArray(i - 1) = Mid(CmdStr, i, 1)
            dataArray(i - 1) = Asc(CharArray(i - 1))
        Next i
        dataArray(StringLength) = 13 ' carriage return
        ' dataArray(StringLengthPlus1) = 10 ' carriage return
    End If
    
    TheHdw.Protocol.ports("UART_TX").NWire.MaxHoldUntilTimeout.Value = 0.11
    
    TheHdw.Protocol.ports("UART_TX").NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
    
    TheHdw.Protocol.ports("UART_TX").Modules("UART_read_response_extended").start '***
    
    If UniqueCmdPerSite Then
        For i = 0 To StringLength
            With TheHdw.Protocol.ports("UART_RX").NWire.Frames("UART_Snd")
                .fields("Data_in").Value = PerSiteDataArray(i)
                .Execute
            End With
            TheHdw.Protocol.ports("UART_RX").IdleWait
        Next i
    Else
        For i = 0 To StringLength
            With TheHdw.Protocol.ports("UART_RX").NWire.Frames("UART_Snd")
                .fields("Data_in").Value = dataArray(i)
                .Execute
            End With
            TheHdw.Protocol.ports("UART_RX").IdleWait
        Next i
    End If
    
    Set dspData = TheHdw.Protocol.ports("UART_TX").NWire.CMEM.DSPWave '***
    
    
    
    If UCase(CmdStr) Like UCase("*sc run*") Then
        If True Then
            TheHdw.Wait 0.01
            If hexary <> "" Then
                HexmeasureCurrent = TheHdw.DCVS.Pins(PinTEMP).Meter.Read(tlStrobe, 1000, 30000, tlDCVSMeterReadingFormatAverage)
            End If
            If uvsary <> "" Then
                UvsmeasureCurrent = TheHdw.DCVS.Pins(PinTEMP).Meter.Read(tlStrobe, 1000, 30000, tlDCVSMeterReadingFormatAverage)
            End If
            If hexary <> "" Then
                TheExec.Flow.TestLimit HexmeasureCurrent, ForceVal:=force_val, Tname:=TNTEMP & "_" & UCase(Replace(CmdStr, LCase(" run "), "")), Unit:=unitAmp
            End If
            If uvsary <> "" Then
                TheExec.Flow.TestLimit UvsmeasureCurrent, ForceVal:=force_val, Tname:=TNTEMP & "_" & UCase(Replace(CmdStr, LCase(" run "), "")), Unit:=unitAmp
            End If
        End If
     End If
     
     
    If ExtendedWait > 0.001 Then
        TheHdw.Wait ExtendedWait
    Else
        TheHdw.Wait 0.02
    End If
    TheHdw.Protocol.ports("UART_TX").Enabled = False
    TheHdw.Protocol.ports("UART_RX").Enabled = False
'----------------------------------------add to avoid 2 interations-----------------------
    Dim Prompt_Idx As New SiteLong

             If TheExec.DataManager.instanceName Like "*SC94*" Or TheExec.DataManager.instanceName Like "*SC95*" Or TheExec.DataManager.instanceName Like "*SC96*" Then
                 Prompt_Idx = 0
             Else
                Prompt_Idx = 1
             End If

Dim UseCmdLength As Boolean

                rundsp.ProcessDUTResponse dspData, TResult, Prompt_Idx, StringLength, UseCmdLength

'---------------------------------------------------------------------------------------
    DUTResponse = dspData.Copy
            
    Call LogDUTResponse(DUTResponse, TResult)

    Set SendCmd_CSW = TResult
    
    If CheckPassFail Then
        TheExec.Flow.TestLimit TResult, 1, 1, , , , , , CmdStr, , , , , , , tlForceNone
    End If
'
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function RTOS_RunScenario_MTR_DOE(Optional testName As String, Optional Cmd1 As String, Optional Cmd2 As String, Optional Cmd3 As String, Optional Cmd4 As String, _
            Optional Cmd5 As String, Optional Cmd1TimeOut As Double = 0#, Optional Cmd2TimeOut As Double = 0#, Optional Cmd3TimeOut As Double = 0#, _
            Optional Cmd4TimeOut As Double = 0#, Optional Cmd5TimeOut As Double = 0#, Optional SelsramBit As String) As Long

On Error GoTo errHandler
    
    ReDim GlobalMergeAry(TheExec.sites.Existing.Count) 'for txt data collection
    
'    Dim Cmd1Status As New SiteLong
'    Dim Cmd2Status As New SiteLong
'    Dim Cmd3Status As New SiteLong
'    Dim Cmd4Status As New SiteLong
'    Dim Cmd5Status As New SiteLong
    
    Dim CmdList As String
    Dim CmdListStatus As New SiteLong
    
    Dim CZSetupName As String
    Dim powerPin As String
    Dim SupplyVoltage As Long
    Dim TRef_Before_Pat As Double               '<- Code timing
    Dim TExec_Before_Pat As Double              '<- Execution time
    Shmoo_Pattern = testName
    Dim LogTimes As Boolean
   
'    LogTimes = True
'    If (LogTimes = True) Then
'        TRef_Before_Pat = TheExec.Timer(0)
'    End If
    
    Dim instanceName As String: instanceName = TheExec.DataManager.instanceName
    testName = instanceName
    
    If Cmd1 <> "" Then CmdList = Cmd1
    If Cmd2 <> "" Then CmdList = CmdList + ";" + Cmd2
    If Cmd3 <> "" Then CmdList = CmdList + ";" + Cmd3
    If Cmd4 <> "" Then CmdList = CmdList + ";" + Cmd4
    If Cmd5 <> "" Then CmdList = CmdList + ";" + Cmd5
    
    Dim Check_TestName As Long: Check_TestName = 0
    Dim Key As Variant: Key = Array("CHAR", "HBV")
    Dim Key_Count As Long
    For Key_Count = 0 To UBound(Key)
        Check_TestName = InStr(1, testName, Key(Key_Count)) 'Search the keyword from position 1
        If Check_TestName > 0 Then Check_TestName = Check_TestName + 1
    Next Key_Count
    
    If Check_TestName = 0 Then TheHdw.PinLevels.ApplyPower
        
    If Not (TheExec.DevChar.Setups.IsRunning) Then
        TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain '' use safe voltage to Selsram
    End If
    
    Shmoo_Save_core_power_per_site_for_Vbump ' store voltage into global variable
    
    'Select Sram Start
    Dim uniquesBit As Boolean, site As Variant
    uniquesBit = IsArray(SelsramBit)
    Dim BinCutSelAry() As String
    BinCutSelAry = Split(SelsramBit, ",")
    
    If UBound(BinCutSelAry) = 4 And SelsramBit <> "" Then
        For Each site In TheExec.sites.Selected
            BinCutSelAry(site) = Decide_Switching_Bit_RTOS(BinCutSelAry(site), g_ApplyLevelTimingValt, "RTOS")
        Next site
        SendCmd BinCutSelAry, 0.1
    Else
        SendCmd Decide_Switching_Bit_RTOS("SSSS", g_ApplyLevelTimingValt, "RTOS"), 0.1
    End If
    
    '//Change to Valt mode
    TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt '' only change the corepower to alt
    TheHdw.Wait 0.005
    'Scenario Run Conditions
    Dim CP_GB_Record As Double
    Dim Bincut_PCPU_value As New SiteDouble
    Dim Bincut_ECPU_value As New SiteDouble
    Dim Bincut_GPU_value As New SiteDouble
    Dim Bincut_PCPU_Minus_GB As New SiteDouble
    Dim Bincut_ECPU_Minus_GB As New SiteDouble
    Dim Bincut_GPU_Minus_GB As New SiteDouble
    
    
    
    Dim i As Integer
    Dim OffsetRecord As New SiteLong
    Dim MTRString_Tmp(5) As String
    Dim MTRString_ROT(5) As String
    Dim MTRString_ROV(5) As String
    Dim MTRTmpWave As New DSPWave
    Dim MTRTmpWave_2 As New DSPWave
    Dim BeforeMTRTmpWave As New DSPWave
    Dim BeforeMTRTmpWave_2 As New DSPWave
    
    Dim MTRDSPWave_ROT As New DSPWave
    Dim MTRDSPWave_ROV As New DSPWave
    MTRDSPWave_ROT = GetStoredCaptureData("Freq_PCPU4_ROT_85C")
    MTRDSPWave_ROV = GetStoredCaptureData("Freq_PCPU4_ROV_85C")
    MTRTmpWave = GetStoredCaptureData("MTR_POST_TS_P0CPU1_TRIM_85C")
    MTRTmpWave_2 = GetStoredCaptureData("MTR_POST_TS_GFX_TRIM_85C")
    BeforeMTRTmpWave = GetStoredCaptureData("MTR_TS_P0CPU1_TRIM_85C")
    BeforeMTRTmpWave_2 = GetStoredCaptureData("MTR_TS_GFX_TRIM_85C")
    For Each site In TheExec.sites.Active
         MTRString_Tmp(site) = ""
         MTRString_ROT(site) = ""
         MTRString_ROV(site) = ""
    Next site

    OffsetRecord = GetStoredMeasurement("VDD_PCPU_VoltageLevelCNT")
    For Each site In TheExec.sites.Active
        For i = 16 - OffsetRecord(site) To 16
            MTRString_ROT(site) = MTRString_ROT(site) + CStr(Math.Round(MTRDSPWave_ROT.Element(i), 8))
            MTRString_ROV(site) = MTRString_ROV(site) + CStr(Math.Round(MTRDSPWave_ROV.Element(i), 8))
            MTRString_ROT(site) = MTRString_ROT(site) + " "
            MTRString_ROV(site) = MTRString_ROV(site) + " "
        Next i
        MTRString_ROT(site) = "mtr cal-rot Tp4i " + MTRString_ROT(site)
        MTRString_ROV(site) = "mtr cal-rov Tp4i " + MTRString_ROV(site)
        MTRString_Tmp(site) = "mtr cal-t Tp4i " + CStr(Math.Round((BeforeMTRTmpWave.Element(0) / 8), 7)) + " " + CStr(Math.Round((MTRTmpWave.Element(0) / 8), 7))
    Next site

    '-----------------------------------MTR Calibration ------------
'     For Each site In theexec.sites
'       CurrentPassBinCutNum(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
'       CP_GB_Record = BinCut(VddBinStr2Enum("VDD_PCPU_MC601"), CurrentPassBinCutNum).FT2_GB(0)
'       Bincut_PCPU_value(site) = UDRP_Fuse.Category(UDRP_Index("VDD_PCPU_MC601")).Read.Value(site)
'       Bincut_PCPU_Minus_GB(site) = Bincut_PCPU_value - CP_GB_Record
'       theexec.DataLog.WriteComment "Site(" & site & ")  VDD POCPU Lowest LVCC : " & Bincut_PCPU_Minus_GB(site) & "mV"
'     Next site
     SendCmd "pmgr mode -e 4", 0.3
     SendCmd MTRString_ROT, 0.6
     SendCmd MTRString_ROV, 0.3
     SendCmd MTRString_Tmp, 0.3
     
     
     
     
     
    MTRDSPWave_ROT = GetStoredCaptureData("Freq_GPU2_ROT_85C")
    MTRDSPWave_ROV = GetStoredCaptureData("Freq_GPU2_ROV_85C")
'    MTRTmpWave = GetStoredCaptureData("MTR_TS_P0CPU1_TRIM_85C")
'    MTRTmpWave_2 = GetStoredCaptureData("MTR_TS_GFX_TRIM_85C")

    For Each site In TheExec.sites.Active
         MTRString_Tmp(site) = ""
         MTRString_ROT(site) = ""
         MTRString_ROV(site) = ""
    Next site

    OffsetRecord = GetStoredMeasurement("VDD_GPU_VoltageLevelCNT")
    For Each site In TheExec.sites.Active
        For i = 16 - OffsetRecord(site) To 16
            MTRString_ROT(site) = MTRString_ROT(site) + CStr(Math.Round(MTRDSPWave_ROT.Element(i), 8))
            MTRString_ROV(site) = MTRString_ROV(site) + CStr(Math.Round(MTRDSPWave_ROV.Element(i), 8))
            MTRString_ROT(site) = MTRString_ROT(site) + " "
            MTRString_ROV(site) = MTRString_ROV(site) + " "
        Next i
        MTRString_ROT(site) = "mtr cal-rot Tg2i " + MTRString_ROT(site)
        MTRString_ROV(site) = "mtr cal-rov Tg2i " + MTRString_ROV(site)
        MTRString_Tmp(site) = "mtr cal-t Tg2i " + CStr(Math.Round((BeforeMTRTmpWave_2.Element(0) / 8), 7)) + " " + CStr(Math.Round((MTRTmpWave_2.Element(0) / 8), 7))
    Next site
     SendCmd MTRString_ROT, 0.6
     SendCmd MTRString_ROV, 0.3
     SendCmd MTRString_Tmp, 0.3
     
    '-----------------------------------MTR Calibration ------------
    CmdListStatus = 0
    
    If CmdList <> "" Then Set CmdListStatus = SendCmd(CmdList, Cmd1TimeOut, False)

    TheExec.Flow.TestLimit CmdListStatus, 1, 1       ', , , , , , TestName
    
    'For Each site In TheExec.sites.Selected
    
            
            '============mtr efuse doe================================
            TheExec.Datalog.WriteComment ""
            TheExec.Datalog.WriteComment "****** MTR RTOS to EFUSE Hex2Decimal*******"
                    Dim rtos_mtr_fuse_name(43) As String
                    Dim mm As Integer
                    Dim rtos_mtr_fuse_value(43) As SiteLong
                    For mm = 0 To 43
                        If mm <= 20 Then
                          rtos_mtr_fuse_name(mm) = "mtr_sense_vt_tp4i_poly_c" & Trim(Str(mm)) & "="
                        ElseIf mm = 21 Then
                          rtos_mtr_fuse_name(mm) = "mtr_polynom_scaler_tp4i" & "="
                        ElseIf mm >= 22 And mm <> 43 Then
                          rtos_mtr_fuse_name(mm) = "mtr_sense_vt_tg2i_poly_c" & Trim(Str(mm - 22)) & "="
                        ElseIf mm = 43 Then
                          rtos_mtr_fuse_name(mm) = "mtr_polynom_scaler_tg2i" & "="
                        End If
                    Next mm
                    Dim strlen As Long
                    Dim fuse_name_len As Long
                    Dim fuse_code_idx As Long
                    Dim fuse_code_value As String
                    Dim fuse_start_lo As Long
                    Dim fuse_name_in_cate As String
                    For Each site In TheExec.sites.Selected
                      If CmdListStatus(site) = 1 Then
                            strlen = Len(GlobalMergeAry(site))
                            For mm = 0 To 43
                                If InStr(1, GlobalMergeAry(site), rtos_mtr_fuse_name(mm)) <> 0 Then
                                   fuse_name_len = Len(rtos_mtr_fuse_name(mm))
                                   fuse_code_idx = InStr(1, GlobalMergeAry(site), rtos_mtr_fuse_name(mm))
                                   fuse_start_lo = fuse_code_idx + fuse_name_len
                                   fuse_name_in_cate = Replace(rtos_mtr_fuse_name(mm), "=", "")
                                   fuse_code_value = Mid(GlobalMergeAry(site), fuse_start_lo, 10) 'Mid(GlobalMergeAry(site), i, 1);
                                   TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(site) + ") " + " MRT Hex EFUSE from UART              " + rtos_mtr_fuse_name(mm) + fuse_code_value
                                   Call auto_eFuse_SetWriteDecimal("CFG", fuse_name_in_cate, fuse_code_value, True)
                                   TheExec.Datalog.WriteComment ""
                                   
                                Else
                                  Exit Function
                                End If
                            Next mm
                       End If
                    Next site
             
                 
            '============mtr efuse doe end===========================
            
     'Next site
    RTOS_UART_Print instanceName, CmdListStatus
    
    
'    For Each site In TheExec.sites.Active
'        If (LogTimes = True) Then
'            TExec_Before_Pat = TheExec.Timer(TRef_Before_Pat)
'            TheExec.DataLog.WriteComment "ElapsedTime Pat Site (" & site & ")" + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
'        End If
'    Next site

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function RTOS_Voltage_Rampdown()
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    Dim step As Integer
    Dim StepNum As Integer
    On Error GoTo errHandler
    
    StepNum = g_RTOSRampStep
    
    TheExec.DataManager.DecomposePinList "CorePower", p_ary, p_cnt

    For step = 1 To StepNum
        For i = 0 To p_cnt - 1
            If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
               If step Mod 2 = 1 Then
                  For Each site In TheExec.sites
                      TheHdw.DCVS.Pins(p_ary(i)).Voltage.Alt.Value = g_ApplyLevelTimingVmain.Pins(p_ary(i)).Value - ((g_ApplyLevelTimingVmain.Pins(p_ary(i)).Value - g_RTOS_SceVoltage.Pins(p_ary(i)).Value) / StepNum) * step
                  Next site
               ElseIf step Mod 2 = 0 Then
                  For Each site In TheExec.sites
                      TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = g_ApplyLevelTimingVmain.Pins(p_ary(i)).Value - ((g_ApplyLevelTimingVmain.Pins(p_ary(i)).Value - g_RTOS_SceVoltage.Pins(p_ary(i)).Value) / StepNum) * step
                  Next site
               End If
            End If
        Next i
        If step Mod 2 = 1 Then
           TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt
           TheHdw.Wait 20 * 0.000001
        ElseIf step Mod 2 = 0 Then
           TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
           TheHdw.Wait 20 * 0.000001
        End If
    Next step
   Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> Shmoo_Save_core_power_per_site:: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function RTOS_Voltage_RampUp()
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    Dim step As Integer
    Dim StepNum As Integer
    
    On Error GoTo errHandler
    

    StepNum = g_RTOSRampStep
        
    TheExec.DataManager.DecomposePinList "CorePower", p_ary, p_cnt
    
    For step = 1 To StepNum
        For i = 0 To p_cnt - 1
            If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
               If step Mod 2 = 1 Then
                  For Each site In TheExec.sites
                      TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = g_RTOS_SceVoltage.Pins(p_ary(i)).Value + ((g_ApplyLevelTimingVmain.Pins(p_ary(i)).Value - g_RTOS_SceVoltage.Pins(p_ary(i)).Value) / StepNum) * step
                  Next site
               ElseIf step Mod 2 = 0 Then
                  For Each site In TheExec.sites
                      TheHdw.DCVS.Pins(p_ary(i)).Voltage.Alt.Value = g_RTOS_SceVoltage.Pins(p_ary(i)).Value + ((g_ApplyLevelTimingVmain.Pins(p_ary(i)).Value - g_RTOS_SceVoltage.Pins(p_ary(i)).Value) / StepNum) * step
                  Next site
               End If
            End If
        Next i
        If step Mod 2 = 1 Then
           TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
           TheHdw.Wait 20 * 0.000001
        ElseIf step Mod 2 = 0 Then
           TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt
           TheHdw.Wait 20 * 0.000001
        End If
    Next step
   Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> Shmoo_Save_core_power_per_site:: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function Decide_Switching_Bit_RTOS(digSrc_EQ As String, Optional DC_Level As PinListData, Optional BlockType As String, Optional shmoo_pin As String, Optional ShmooPinsVoltage As PinListData, Optional ForcePin As String, Optional SetForceVoltage As Dictionary, Optional BincutEnable As Boolean = False) As String

  Dim site As Variant
  Dim logicPin As String
  Dim SramPin As String
  Dim DSSC_Switching_Voltage As New PinListData
  Dim Sdomain() As Long
  Dim SramValue As Double
  Dim DSSCSelSrmOpposite As Long
  Dim BlockTypeNum As Long
  Dim i As Integer, j As Integer
  Dim ReturnString() As String
  
  On Error GoTo errHandler
  BlockTypeNum = -1
  ReDim ReturnString(Len(digSrc_EQ) - 1)
  If BincutEnable = False Then
    Decide_DSSC_Switching_Voltage DSSC_Switching_Voltage, DC_Level, shmoo_pin, ShmooPinsVoltage, ForcePin, SetForceVoltage
  End If
    '///find blocktype
  
 For i = 0 To UBound(SelsramMapping)
    If UCase(SelsramMapping(i).blockName) <> "*" Then
      If UCase(BlockType) Like "*" & UCase(SelsramMapping(i).blockName) & "*" Then
         BlockTypeNum = i
         Exit For
      End If
    End If
 Next i


  Dim p_ary() As String, p_cnt As Long
  Dim LogicValue() As Double
  
 
  If BlockTypeNum <> -1 Then
    For i = 0 To Len(digSrc_EQ) - 1
        If UCase(CStr(Mid(digSrc_EQ, i + 1, 1))) Like "S" Then
        
            logicPin = SelsramMapping(BlockTypeNum).logic_Pin(i)
            SramPin = SelsramMapping(BlockTypeNum).sram_Pin(i)
            DSSCSelSrmOpposite = SelsramMapping(BlockTypeNum).SelSrm1(i)
             
            If UCase(logicPin) Like "PRESERVED" Then
                ReturnString(i) = DSSCSelSrmOpposite
            Else
                TheExec.DataManager.DecomposePinList logicPin, p_ary, p_cnt
                For j = 0 To p_cnt - 1
                    ReDim Preserve LogicValue(j)
                    LogicValue(j) = CDbl(DSSC_Switching_Voltage.Pins(p_ary(j)).Value)
                Next j
                SramValue = CDbl(DSSC_Switching_Voltage.Pins(SramPin).Value)
                If DSSCSelSrmOpposite = 0 Then
                    ReDim Preserve Sdomain(UBound(LogicValue))
                    For j = 0 To UBound(LogicValue)
                        If j = 0 Then
                            Sdomain(j) = IIf((LogicValue(j) > SramValue), 1, 0)
                        Else
                            Sdomain(j) = IIf((LogicValue(j) > SramValue), 1, 0)
                            If Not Sdomain(j) = Sdomain(j - 1) Then
                               TheExec.ErrorLogMessage "PinGroup with different SELSRM value"
                            End If
                        End If
                    Next j
                    ReturnString(i) = Sdomain(LBound(Sdomain))
                ElseIf DSSCSelSrmOpposite = 1 Then
                    ReDim Preserve Sdomain(UBound(LogicValue))
                    For j = 0 To UBound(LogicValue)
                        If j = 0 Then
                            Sdomain(j) = IIf((LogicValue(j) > SramValue), 0, 1)
                        Else
                            Sdomain(j) = IIf((LogicValue(j) > SramValue), 0, 1)
                            If Not Sdomain(j) = Sdomain(j - 1) Then
                                TheExec.ErrorLogMessage "PinGroup with different SELSRM value"
                            End If
                        End If
                    Next j
                    ReturnString(i) = Sdomain(LBound(Sdomain))
                End If
            End If
        Else
            ReturnString(i) = CDbl(Mid(digSrc_EQ, i + 1, 1))
        End If
    Next i
    
    For i = 0 To UBound(ReturnString)
        ReturnString(i) = SelsramMapping(BlockTypeNum).comment(i) & Replace(Replace(ReturnString(i), "1", "ff"), "0", "00")
    Next i
      
    Decide_Switching_Bit_RTOS = Join(ReturnString, ";")
    
    g_RTOS_SceVoltage = DSSC_Switching_Voltage.Copy
    
  End If
  
  Exit Function
  
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + "Decide_Switching_Bit" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function RTOS_Boot_Up_fail_Power_Up()
'''    '======================Relay settings for power up======================= ' 200205
'''        Dim Relay_on_Conti As New PinList
'''        Dim Relay_off_Conti As New PinList
'''        Dim Relay_on_Default As New PinList
'''        Dim Relay_off_Default As New PinList
'''        Relay_on_Conti.Value = "K06,K08,K10,K12,K14,K16,K18,K20,K22,K24,K26,K28,K30,K42,K44,K46,K48,K49,K50,K51,K52,K55,K56,K59,K61,K63,K65,K67,K69,K71"
'''        Relay_off_Conti.Value = "K01,K02,K03,K04,K05,K07,K09,K11,K13,K15,K17,K19,K21,K23,K25,K27,K29,K31,K32,K33,K34,K35,K36,K37,K38,K39,K40,K41,K43,K45,K47,K53,K54,K57,K58,K60,K62,K64,K66,K68,K70"
'''        Relay_on_Default.Value = "K06,K08,K10,K12,K14,K16,K18,K20,K22,K24,K26,K28,K30,K42,K44,K46,K61,K63,K65,K67,K69,K71"
'''        Relay_off_Default.Value = "K01,K02,K03,K04,K05,K07,K09,K11,K13,K15,K17,K19,K21,K23,K25,K27,K29,K31,K32,K33,K34,K35,K36,K37,K38,K39,K40,K41,K43,K45,K47,K48,K49,K50,K51,K52,K53,K54,K55,K56,K57,K58,K59,K60,K62,K64,K66,K68,K70"
'''    '======================Relay settings for power up======================= ' 200205
'''
'''    '======================If Shmoo fail,do power up again======================= ' 200205
''''        TheExec.DataManager.DecomposePinList relay_on, Pins_On(), Pin_Cnt_On
''''        TheExec.DataManager.DecomposePinList relay_off, Pins_Off(), Pin_Cnt_Off
'''        Relay_Control Relay_on_Conti, Relay_off_Conti, 0.003
'''        PowerUp_Parallel "All_Power", "", "", "", "", "All_Digital_PowerUp", 0.001, -1
'''        Relay_Control Relay_on_Default, Relay_off_Default, 0.003
'''    '======================If Shmoo fail,do power up again======================= ' 200205
End Function
