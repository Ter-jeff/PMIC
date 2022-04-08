Attribute VB_Name = "Exec_IP_Module"
Option Explicit
'Public write_SPIROM_CheckSum As Integer

Public Const Exec_IP_Module = "0.0" ''lib version, initial version for central library.
Public write_spirom As New SiteBoolean
Public Mbist_Repair_CompareType As Variant  'for Mbist finger print
Public glb_TesterType As String

' This module contains empty Exec Interpose functions (see online help
' for details).  These are here for convenience and are completely optional.
' It is not necessary to delete them if they are not being used, nor is it
' necessary that they exist in the program.



' Immediately at the conclusion of the initialization process.
' Do not program test system hardware from this function.
Function OnTesterInitialized()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    ' OnTesterInitialized executes before TheExec is even established so nothing
    ' better to do then msgbox in this case.  Note that unhandled errors can allow the
    ' user to press "End" which will result in a DataTool crash.  Errors in this routine
    ' need to be debugged carefully.
    MsgBox "Error encountered in Exec Interpose Function OnTesterInitialized" + vbCrLf + _
        "VBT Error # " + Trim(Str(err.number)) + ": " + err.Description
                If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately at the conclusion of the load process.
' Do not program test system hardware from this function.
Function OnProgramLoaded()

    On Error GoTo errHandler

    ' Put code here
    #If IGXL8p30 Then
    #Else
    TheHdw.Digital.LevelSets.OptimizeAllocation = True      'add to avoid levelsets over 255 limitation 2018/01/09
    Call tl_activateallusersheets
        #End If

    ' Put code here
    'for TERA1 encryption need, add following code
    If Not TheExec.SoftwareVersion Like "8.10.90_uflx*" Then
        m_cpcmodule.SuppressCheckForUnProtectedPatterns = True
    End If
    
    'for 8.30 encryption need, add following code
    m_STDSvcClient.CPCModule.SuppressCheckForUnProtectedPatterns = True
    
    If TheExec.SoftwareVersion Like "*9.10*" Then
        CallByName TheExec.TestProgram, "MemoryLimitCheckEnabled", VbLet, False
    End If
    
    'for relax reference clock to lower frequency
    'Enable the full frequency range for the nWire PA clock
''    TheHdw.Digital.Timing.FullPAClockFrequencyRange = True
''
    'turnning on the simulator
    'note if so, the test time of simulation would increase
    'TheExec.Simulator.ForceAllSimulation (tlSimForce)
    
    TheExec.DataManager.MaxSheetValidationErrorEnabled = False
    TheHdw.Digital.EnablePinRespecification = True
    
    If is_reference_installed("Scripting") = False Then
        Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\WINDOWS\system32\scrrun.dll"
    End If
    
     If is_reference_installed("VBScript_RegExp_55") = False Then
        Application.VBE.ActiveVBProject.References.AddFromFile "C:\WINDOWS\system32\vbscript.dll\3"
    End If
    
    
  '  If is_reference_installed("PATTERNDATAMANAGERLib") = False Then
  '      If Application.OperatingSystem Like "*NT 6*" Then
  '          Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\Program Files (x86)\Teradyne\IG-XL\8.10.12_uflx\bin\PatternDataManager.dll"
  '      Else
  '          Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\Program Files\Teradyne\IG-XL\8.10.90_uflx\bin\PatternDataManager.dll"
  '      End If
  '  End If
  '
   'nWire PA for XI0
   Call TheHdw.Protocol.Families("nWire").Types.Add(".\xml_Files\FreeRunClk_TDR_TRUE_32Clk_8Idle.xml", "Clock")
   Call TheHdw.Protocol.Families("nWire").Types.Add(".\xml_Files\FreeRunClk_differential.xml", "Clock_Diff")
'   Call TheHdw.Protocol.Families("FRC").Types.Add(".\xml_Files\FRC.pa")
'   Call TheHdw.Protocol.Families("FRC").Types.Add(".\xml_Files\FRCRef.pa")
'   Call TheHdw.Protocol.Families("FRC").Types.Add(".\xml_Files\FRCRef_differential.pa", "FRC_Clock_Diff")
    TheHdw.Digital.EnableSharedsiteSupportCheck = True

    'nWire PA for UART receiver need to take care if other pattern use this PA pin.
    'Call TheHdw.Protocol.Families("nWire").Types.Add(".\xml_Files\UART_x3_RX.xml", "UART") ' 20160322
    Call TheHdw.Protocol.Families("nWire").Types.Add(".\xml_Files\UART_x3_RX.xml", "UART_PA_RX") ' 20160322
    Call TheHdw.Protocol.Families("nWire").Types.Add(".\xml_Files\UART_x3_TX.xml", "UART_PA_TX") ' 20160603 Leslie

    '''''''''20180628 add to prevent shmoo can't read sheet
    Dim CZ_Activate_Sheet As Worksheet
    For Each CZ_Activate_Sheet In ThisWorkbook.Sheets
            If UCase(CZ_Activate_Sheet.Name) Like "*FLOW_DCTEST*" Or UCase(CZ_Activate_Sheet.Name) Like "*FLOW_HARDIP*" Then
                CZ_Activate_Sheet.Activate
            End If
    Next CZ_Activate_Sheet
    ''''''''
    

    Exit Function
errHandler:
    HandleExecIPError "OnProgramLoaded"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation succeeds.
Function OnProgramValidated()
    On Error GoTo errHandler
    
    '''20200309: Modified to re-do initVddBinning when BinCut testjob or sheet is changed.
    FlagInitVddBinningTable = False
    
   ' write_SPIROM_CheckSum = 0
    
    ''''flag initilize
    Flag_RAK_INIT = False   'Not initialized, Init in HardIP flow
    Flag_RSCR_INIT = False  'Not initialized, Init in MBist flow
    Flag_Shmoo_INIT = False 'Not initialized, Init in Shmoo flow
    Flag_MBISTFailBlock_INIT = False

    TheHdw.Patterns.EnableExplicitFileNames = True  'to fix dssc can not recognize pat.gz format
    TheExec.Flow.HighParallelMode = True            '140501 pre-shut down in parallel, TTR purpose, false to check if nWire not stop by site fail
    Ignore_nWire_Error                              '140328 2D-shmoo nWire debug
    
    Init_Datalog_Setup          'datalog settings for the requirement of product
    
    TheHdw.Digital.Pins("Cal_Excluded").Calibration.Excluded = True 'bypass TDR
    
    'CharStoreResultsUntilNextRun, clear shmoo momory to prevent crash
    If LCase(TheExec.CurrentJob) Like "char*" Or TheExec.Flow.EnableWord("Shmoo_BringUp") Then
        TheExec.DevChar.Configuration.Features.Item(tlDevCharFeature_StoreResultsUntilNextRun).Enabled = False
        m_STDSvcClient.SelfTest.MemoryCollectRunInterval = 1
    End If

 '   TheExec.Flow.EnableWord("Read_EEPROM_DIBID") = True 'for DIB Board ID read out
    
    '*** protect efuse sheets ***
'    gL_1st_FuseSheetRead = 0 ''''20150624
'    Call UnProtect_eFuse_Sheet
'    Call autoArrange
'    Call autoArrange("UDR_compare_ChkList_appA")
'    Call Protect_eFuse_Sheet ''''MUST be here
    
    ''''---------Start of Mbist ChkList---------------
'    gL_1st_MbistSheetRead = 0 ''''20151020
    ''Call UnProtect_Mbist_Sheet
    ''Call Protect_Mbist_Sheet
    ''''---------  End of Mbist ChkList---------------
     
     ' 20150121 for SPI ROM auto fuse
    TheExec.Flow.EnableWord("Write_SPIROM") = True 'trigger SPI ROM write
    
    ' 20150128 - Load PA UART RX module
        'only enable it during SPI debug.
    'TheHdw.Protocol.Ports("UART_PA").ModuleFiles.UnloadAll
    'Call PreLoad_PA_Modules("READ", 15000, "UART_PA")
    'only enable it during SPI debug.
    
    
    write_spirom = True '20160324 central lib

    ' 20160422 - Write Bin Name to STDF
    Call Bintable_initial
    TheHdw.Digital.Alarm(tlHSDMAlarmAll) = tlAlarmForceBin
    
    Call ProcessAssignment ' Added By Oscar 180523 For Read DigSrc assignment from Sheet(Function is in LIB_HardIP)
'    Call HardIP_OnProgramStarted_Process
    
    power_up_en = False 'initial for switching job
    
     ''' to expand fail cycle number more than 100 million (from 10 million) - 190408
    TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Cycle.Width = 9
    TheExec.Datalog.ApplySetup
    
        ' for License Check, initial global variable
        gL_License_check = 0
        
    Exit Function
errHandler:
    HandleExecIPError "OnProgramValidated"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation fails.
Function OnProgramFailedValidation()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramFailedValidation"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately at the conclusion of the user DIB calibration process (previously
' known as the TDR calibration process). Called only if user DIB calibration succeeds.
Function OnTDRCalibrated()
    On Error GoTo errHandler

    ' Put code here
    'To fix freerunclock unstable issue used for  8.10.12 is fixed in 8.30
    'm_stdsvcclient.TesterSupport.TimingCalChannel.DriveEnable
    
    Exit Function
errHandler:
    HandleExecIPError "OnTDRCalibrated"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately after "pre-job reset" when the test program starts.
' Note that "first run" actions can be enclosed in
' If TheExec.ExecutionCount = 0 Then...
' (see online help for ExecutionCount)
Function OnProgramStarted()
    On Error GoTo errHandler
    Dim i As Integer, j As Integer
    
    glb_TesterType = TheHdw.Tester.Type
    
    currentJobName = LCase(TheExec.CurrentJob) ''Carter, 20191115
    ' Put code here
    Call RemoveAllStored
    gl_IDS_INFO_Dic.RemoveAll ''Carter, 20191115
    'Add for RAK
    Call HardIP_OnProgramStarted_Process
  
    'use "Parse_SELSRM_Mapping_Table" for parsing selsrm table
    'Call Parsing_SELSRM_Mapping_Table '"Parsing" is for Bincut, we call it in initVDDbinning
    Call Parse_SELSRM_Mapping_Table '"Parse" is for CHAR/function_T
    Call Parse_EMA_DigSrcInfo ''Carter, 20191115
    Call initVddBinning

    ''Export UnExistPins
    If TheExec.EnableWord("aExportUnExistPins") = True Then
        Call Search_UnExistPin
    End If
    
    TheHdw.Digital.CheckContextExclusion = True 'nWire context check flag, need to think about how to switch between engineering mode and production mode.
    
    Init_DIB_Power  'reset DIB power
    
    currentJobName = LCase(TheExec.CurrentJob)  'init job string, reset every program start
    gL_ProductionTemp = ""  'init temparature
    TheExec.Flow.FlowFlagMode = tlFlowFlagLatchTestResult 'prevent next pass overwrite previous  fail
    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug   'use host computer to cal dsp function
    
        'Get differential pair
    Call RetrieveDictionaryOfDiffPairs

    '*** eFuse initial
    ''''20150624 move to Flow_Table_Main_Init_Flows
''    Call auto_eFuse_Initialize
    
    '' 20151029 - Compensate DIB impedence for RAK
    Dim ws_def As Worksheet
    Dim wb As Workbook
    Dim siteIdx As Integer
    Dim SiteNum1 As Integer
    Dim SiteNum2 As Integer
    Dim DACInitialFlag As Boolean

    Set wb = Application.ActiveWorkbook

    If TheExec.Flow.EnableWord("DebugPrintFlag") = True Then
        DebugPrintFlag_Chk = True
    Else
        DebugPrintFlag_Chk = False
    End If

    DACInitialFlag = False  '''added for NonAP flag initial

    Find_nWire_Pin   '''update for multiple nWire CLK, 2017/07/18
''
Dim site As Variant
For Each site In TheExec.sites
    If write_spirom = True Then
        TheExec.Flow.EnableWord("Write_SPIROM") = True
    End If

Next site
''
    gl_UseStandardTestName_Flag = True
    If gl_UseStandardTestName_Flag Then
        Call SetupDatalogFormat(TestNameW:=90, PatternW:=100)
    End If
    
    If TheExec.Flow.EnableWord("production") = True Then
        TheExec.Datalog.WriteComment "[HIP TTR EnableWord: Production]"
    ElseIf TheExec.Flow.EnableWord("monitoring") = True Then
        TheExec.Datalog.WriteComment "[HIP TTR EnableWord: Monitoring]"
    ElseIf TheExec.Flow.EnableWord("char") = True Then
        TheExec.Datalog.WriteComment "[HIP TTR EnableWord: Char]"
    Else
        TheExec.Datalog.WriteComment "[HIP TTR EnableWord: None]"
    End If
    
    ''''' 20180710 Add initialize value ''''''''''''
    CHAR_USL_HVCC = 9999
    CHAR_USL_LVCC = 9999
    CHAR_LSL_HVCC = 9999
    CHAR_LSL_LVCC = 9999

    'for Mbist finger print
    Mbist_Repair_CompareType = "Cycle"
    If TheExec.Flow.EnableWord("Mbist_FingerPrint_Vector") = True Then Mbist_Repair_CompareType = "Vector"
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramStarted"
    If AbortTest Then Exit Function Else Resume Next
End Function
 

' Immediately before "post-job reset" when the test program completes.
' Note that any actions taken here with respect to modification of binning
' will affect the binning sent to the Operator Interface, but will not affect
' the binning reported in Datalog.
Function OnProgramEnded()
    On Error GoTo errHandler

    ' Put code here
    TheHdw.DIB.powerOn = False  'reset DIB power, prevent hot switch


    Exit Function
errHandler:
    HandleExecIPError "OnProgramEnded"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately before a site is disconnected.
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
Function OnPreShutDownSite()
    On Error GoTo errHandler

    ' Put code here
    If power_dcvs_exit = True Or power_dcvi_exit = True Then
        PowerDown_Parallel AllPowerPinlist, AllDCVIPinlist, All_DigitalPinlist_Disc, , True ' powerdown once bin out
    End If

    Exit Function
errHandler:
    HandleExecIPError "OnPreShutDownSite"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
' Immediately after a site is disconnected.
Function OnPostShutDownSite()
    On Error GoTo errHandler

    ' Put code here
    

    Exit Function
errHandler:
    HandleExecIPError "OnPostShutDownSite"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately befoe any new calibration factors are loaded
' or new calibrations run.  Not called if no action is taken during AutoCal.
Function OnAutoCalStarted()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnAutoCalStarted"
    If AbortTest Then Exit Function Else Resume Next
End Function

' Immediately after AutoCal has completed.
' Not called no action has been taken (new factors loaded, or cal performed).
Function OnAutoCalCompleted()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnAutoCalCompleted"
    If AbortTest Then Exit Function Else Resume Next
End Function


' Called right before an alarm is reported
' The alarmList is a tab delimited string of alarm error messages
Function OnAlarmOccurred(alarmList As String)

    On Error GoTo errHandler
    
'    UNCOMMENT TO THE FOLLOWING LINES TO PARSE ALARMS

    Dim i As Long
    Dim alarmArray() As String
    Dim s As Long
    ' The string is a tab delimited list of alarm error messages
    alarmArray = Split(alarmList, vbTab)

    ' This will loop through all the alarms
    For i = LBound(alarmArray) To UBound(alarmArray)
        ' Then you can print it
'        Debug.Print "Alarm " & i & ": " & alarmArray(i)

        ' Or check for a specific error
'        If InStr(1, alarmArray(i), "DCVS:0001") Then
'            Debug.Print "Found DCVS Alarm 1!!"
'        End If

                If InStr(1, alarmArray(i), "Site ") <> 0 Then   '''' We need to add
            s = CLng(Mid(alarmArray(i), InStr(1, alarmArray(i), "Site ") + 5, 1))
            If s >= 0 Then
                alarmFail(s) = True
            End If
        End If  '''' We need to add

    Next i

                
    Exit Function
errHandler:
    HandleExecIPError "OnAlarmOccurred"
    If AbortTest Then Exit Function Else Resume Next
End Function

' When the user pressed the VB Stop button, this interpose function would be called after OnPostShutDownSite was called.
' The user would put code here to make sure global variable are created and contain the correct data.
Function OnGlobalVariableReset()
    On Error GoTo errHandler

    ' Put code here
    Call ProcessAssignment ' Added By Oscar 180523 For Read DigSrc assignment from Sheet(Function is in LIB_HardIP)
    'Call HardIP_OnProgramStarted_Process
    
    Exit Function
errHandler:
    HandleExecIPError "OnGlobalVariableReset"
    If AbortTest Then Exit Function Else Resume Next
End Function

' Immediately once Vaildation get started
Function OnValidationStart()
    On Error GoTo errHandler

    ' Put code here

    Exit Function
errHandler:
    HandleExecIPError "OnValidationStart"
    If AbortTest Then Exit Function Else Resume Next
End Function
' Immediately at the conclusion of the workbook close process. The function is called in any of the following options,
' File->Close
' File->Exit
' Directly triggered the close (?X?) button of the workbook.
Function OnProgramClose()
    On Error GoTo errHandler

    ' Put code here
    TheHdw.DIB.powerOn = False  'reset DIB power


    Exit Function
errHandler:

    HandleExecIPError "OnProgramClose"
    If AbortTest Then Exit Function Else Resume Next

End Function




Public Function Init_Datalog_Setup()    'in on program validated
    
    TheExec.Datalog.Setup.DatalogSetup.PartResult = True
    TheExec.Datalog.Setup.DatalogSetup.XYCoordinates = True
    TheExec.Datalog.Setup.DatalogSetup.DisableChannelNumberInPTR = True 'disable channel name to stdf, PE's datalog request -- 131225, chihome
    TheExec.Datalog.Setup.DatalogSetup.OutputWidth = 0
    If LCase(TheExec.CurrentJob) Like "*char*" Then
        TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = 75
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.testName.Width = 60
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 70
    End If
    TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
End Function
Public Function Ignore_nWire_Error()
    TheExec.Error.Behavior("HSDMPI:1335") = tlErrorIgnore
    TheExec.Error.Behavior("HSDMPI:0109") = tlErrorIgnore
    TheExec.Error.Behavior("NWirePI:0074") = tlErrorIgnore
End Function

Public Function Init_DIB_Power()
    TheHdw.DIB.powerOn = False  'reset DIB power
    TheHdw.DIB.Power.Item("12V").State = tlOn
    TheHdw.DIB.Power.Item("5V_1").State = tlOn
    TheHdw.DIB.Power.Item("5V_2").State = tlOn
    TheHdw.DIB.Power.Item("3.3V").State = tlOn
End Function
