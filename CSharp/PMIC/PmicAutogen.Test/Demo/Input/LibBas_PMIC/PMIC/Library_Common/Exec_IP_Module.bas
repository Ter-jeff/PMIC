Attribute VB_Name = "Exec_IP_Module"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

Public Const Exec_IP_Module = "0.0"    ''lib version, initial version for central library.
''Public write_SPIROM_CheckSum As Integer
''Public write_spirom As New SiteBoolean

' This module contains empty Exec Interpose functions (see online help
' for details).  These are here for convenience and are completely optional.
' It is not necessary to delete them if they are not being used, nor is it
' necessary that they exist in the program.

' Immediately at the conclusion of the initialization process.
' Do not program test system hardware from this function.
Function OnTesterInitialized()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnTesterInitialized"

    ' Put code here


    Exit Function
ErrHandler:
    ' OnTesterInitialized executes before TheExec is even established so nothing
    ' better to do then msgbox in this case.  Note that unhandled errors can allow the
    ' user to press "End" which will result in a DataTool crash.  Errors in this routine
    ' need to be debugged carefully.

    '//2019_1213
    '    MsgBox "Error encountered in Exec Interpose Function OnTesterInitialized" + vbCrLf + _
         '        "VBT Error # " + Trim(Str(err.Number)) + ": " + err.Description
    TheExec.Datalog.WriteComment "Error encountered in Exec Interpose Function OnTesterInitialized" + vbCrLf + _
                                 "VBT Error # " + Trim(Str(err.Number)) + ": " + err.Description
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' Immediately at the conclusion of the load process.
' Do not program test system hardware from this function.
Function OnProgramLoaded()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnProgramLoaded"


    '        If IsModuleExist("VBT_Import_Module") Then
    '                Call Import_Module_Install
    '        End If

    ' Put code here
    TheHdw.Digital.LevelSets.OptimizeAllocation = True      'add to avoid levelsets over 255 limitation 2018/01/09
    Call tl_activateallusersheets    ''''20181019, <IG9.0 MUST>

    'for TERA1 encryption need, add following code
    m_cpcmodule.SuppressCheckForUnProtectedPatterns = True

    'for 8.30 encryption need, add following code
    m_stdsvcclient.CPCModule.SuppressCheckForUnProtectedPatterns = True

    ''''    'for relax reference clock to lower frequency
    ''''    'Enable the full frequency range for the nWire PA clock
    ''''    'TheHdw.Digital.Timing.FullPAClockFrequencyRange = True

    ''''    'turnning on the simulator
    ''''    'note if so, the test time of simulation would increase
    ''''    'TheExec.Simulator.ForceAllSimulation (tlSimForce)

    TheExec.DataManager.MaxSheetValidationErrorEnabled = False
    TheHdw.Digital.EnablePinRespecification = True

    If is_reference_installed("Scripting") = False Then
        Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\WINDOWS\system32\scrrun.dll"
    End If

    If is_reference_installed("VBScript_RegExp_55") = False Then
        ''Application.VBE.ActiveVBProject.References.AddFromFile "C:\WINDOWS\system32\vbscript.dll\3"
        Application.VBE.ActiveVBProject.References.AddFromFile "C:\WINDOWS\system32\vbscript.dll"
    End If

    ''''nWire PA for XI0


    Call TheHdw.Protocol.Families("nWire").Types.Add(".\xml_Files\FreeRunClk_TDR_TRUE_32Clk_8Idle.xml", "Clock")

    ''''nWire PA for SPI, relay maxtix ADG1414
    Call TheHdw.Protocol.Families("nWire").Types.Add(".\xml_Files\SPI.xml", "SPIPORT")
    Call TheHdw.Protocol.Families("nWire").Types.Add(".\xml_Files\nWire_JTAG_BitField.xml", "JTAGPORT")   ' BitField feature
    Call TheHdw.Protocol.Families("nWire").Types.Add(".\xml_Files\SPMIBF.xml", "SPMIPORT")    'BitField feature


    TheHdw.Digital.EnableSharedsiteSupportCheck = True

    ''''    '''''''''20180628 add to prevent shmoo can't read sheet
    ''''    Dim CZ_Activate_Sheet As Worksheet
    ''''    For Each CZ_Activate_Sheet In ThisWorkbook.Sheets
    ''''        If UCase(CZ_Activate_Sheet.Name) Like "*FLOW_DCTEST*" Or UCase(CZ_Activate_Sheet.Name) Like "*FLOW_HARDIP*" Then
    ''''            CZ_Activate_Sheet.Activate
    ''''        End If
    ''''    Next CZ_Activate_Sheet
    ''''    ''''''''

    Exit Function
ErrHandler:
    HandleExecIPError "OnProgramLoaded"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' Immediately at the conclusion of the validate process. Called only if validation succeeds.
Function OnProgramValidated()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnProgramValidated"

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'DO NOT SET SCREEN UPDATING TO FALSE
    'If there is a validation error, the screen becomes stuck
    'To fix this issue we need to address the long validation issues at the root of the problem
    'Please see Duy Lam (duy.lam@teradyne.com) if you have further questions
    'October 21, 2019
    'Application.ScreenUpdating = False


    TheHdw.Patterns.EnableExplicitFileNames = True  'to fix dssc can not recognize pat.gz format
    TheExec.Flow.HighParallelMode = True            '140501 pre-shut down in parallel, TTR purpose, false to check if nWire not stop by site fail
    Ignore_nWire_Error                              '140328 2D-shmoo nWire debug

    Init_Datalog_Setup          'datalog settings for the requirement of product
    Call SetupDatalogFormat(TestNameW:=90, PatternW:=100)

    Call spotcal_Pre_OnProgramValidated
    TheHdw.Digital.EnablePinRespecification = True


    Dim DigSrcPinAry() As String, NumberPins As Long
    Call TheExec.DataManager.DecomposePinList("TDR_Exclude_Pins", DigSrcPinAry(), NumberPins)
    If NumberPins > 0 Then
        TheHdw.Digital.Pins("TDR_Exclude_Pins").Calibration.Excluded = True    'bypass TDR
        TheHdw.Digital.Pins("TDR_Exclude_Pins").Calibration.DIB.Trace = 0.000000065535    'Give default value for non-TDR pins
    End If

    'CharStoreResultsUntilNextRun, clear shmoo momory to prevent crash
    If LCase(TheExec.CurrentJob) Like "char*" Or TheExec.Flow.EnableWord("Shmoo_BringUp") Then
        TheExec.DevChar.Configuration.Features.Item(tlDevCharFeature_StoreResultsUntilNextRun).Enabled = False
        m_stdsvcclient.SelfTest.MemoryCollectRunInterval = 1
    End If

    ' 20150128 - Load PA UART RX module
    'only enable it during SPI debug.
    'TheHdw.Protocol.Ports("UART_PA").ModuleFiles.UnloadAll
    'Call PreLoad_PA_Modules("READ", 15000, "UART_PA")
    'only enable it during SPI debug.
    ''write_spirom = True '20160324 central lib
    ''''could be unused.
    ''power_up_en = False 'initial for switching job

    ' 20160422 - Write Bin Name to STDF
    Call Bintable_initial

    Call Init_DIB_Power  'reset DIB power
    
    Call ReadEEPROM
    TheHdw.Digital.Alarm(tlHSDMAlarmAll) = tlAlarmForceBin

    '''update for multiple nWire CLK, 2017/07/18
    ''Call Find_nWire_Pin  --> need to check is it necessnry?

    ActivateAllSheet

    '***********For OTP we need OTP_register_Map sheet in program***********
    Call ArrangeOtpTable
    '20190718 Reset OTPData structure ( in case there are subtle changes in otp_reg_map )
    g_DictOTPNameIndex.RemoveAll


    '***********For DC table method , we need import table method to program then we can open it***********
    If LeakPinDic.Count = 0 Then
        'GenLeakPinDic        ''' after user apply DC table method then this function could be enable
    End If
    If ContiPinDic.Count = 0 Then
        'GenContiPinDic       ''' after user apply DC table method then this function could be enable
    End If




    '***********For ADG Trace method, we need import relay trace method to program then we can open it***********
    ''MatrixLoc MatrixShtName     ''' after user apply Relay trace method then we could enable this function
    ''GetNameAndNum               ''' after user apply Relay trace method then we could enable this function


    TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = True
    TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = True
    TheExec.Datalog.ApplySetup

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '10172019
    ''''''Fix screen freeze after validation finishes
    'Application.ScreenUpdating = True

    '20200330 JY Need to clear enabled words or the enabled word will be implicitly selected when you change to different 'part'
    TheExec.ClearAllEnableWords
        
    Exit Function
ErrHandler:
    HandleExecIPError "OnProgramValidated"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' Immediately at the conclusion of the validate process. Called only if validation fails.
Function OnProgramFailedValidation()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnProgramFailedValidation"

    ' Put code here

    Exit Function
ErrHandler:
    HandleExecIPError "OnProgramFailedValidation"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' Immediately at the conclusion of the user DIB calibration process (previously
' known as the TDR calibration process). Called only if user DIB calibration succeeds.
Function OnTDRCalibrated()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnTDRCalibrated"

    ' Put code here
    'To fix freerunclock unstable issue used for  8.10.12 is fixed in 8.30
    'm_stdsvcclient.TesterSupport.TimingCalChannel.DriveEnable

    Exit Function
ErrHandler:
    HandleExecIPError "OnTDRCalibrated"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' Immediately after "pre-job reset" when the test program starts.
' Note that "first run" actions can be enclosed in
' If TheExec.ExecutionCount = 0 Then...
' (see online help for ExecutionCount)
Function OnProgramStarted()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnProgramStarted"

    ''''gL_ProductionTemp = ""  'init temparature
    ''''DACInitialFlag = False  '''added for NonAP flag initial
    ''''Call RemoveAllStored
    'Get differential pair
    ''''Call RetrieveDictionaryOfDiffPairs

    ''''' Reading the excel sheet for the OTP dependent VDD of the LDOs. JC

    '''Call ReadBack_LDO_VDD_Table '' LDO dependent on OTP version to control VDD level setting. User need to read table.

    ''' turn on user power supplies
    If TheHdw.DIB.Power.Item("5V_1").State = tlOff Or _
       TheHdw.DIB.Power.Item("5V_2").State = tlOff Or _
       TheHdw.DIB.Power.Item("3.3V").State = tlOff Or _
       TheHdw.DIB.Power.Item("12V").State = tlOff Then
        TheHdw.DIB.Power.Item("5V_1, 5V_2, 3.3V, 12V").State = tlOn
    End If


    TheHdw.Digital.CheckContextExclusion = True    'nWire context check flag, need to think about how to switch between engineering mode and production mode.

    g_sCurrentJobName = LCase(TheExec.CurrentJob)  'init job string, reset every program start
    gL_ProductionTemp = ""  'init temparature
    TheExec.Flow.FlowFlagMode = tlFlowFlagLatchTestResult    'prevent next pass overwrite previous  fail
    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug  'use host computer to cal dsp function
    g_sTPPath = Application.ActiveWorkbook.Path


    If TheExec.Flow.EnableWord("DebugPrintFlag") = True Then
'        DebugPrintFlag_Chk = True                 '- This Global used in DebugPrintFunc for debug print and Functional_T_Update call it.
        gb_DebugPrintFlag_Chk = True
    Else
'        DebugPrintFlag_Chk = False                '- This Global used in DebugPrintFunc for debug print and Functional_T_Update call it.
        gb_DebugPrintFlag_Chk = False
    End If



    gl_UseStandardTestName_Flag = True
    If gl_UseStandardTestName_Flag Then
        Call SetupDatalogFormat(TestNameW:=90, PatternW:=100)
    End If

    ''''' 20180710 Add initialize value ''''''''''''
    CHAR_USL_HVCC = 9999
    CHAR_USL_LVCC = 9999
    CHAR_LSL_HVCC = 9999
    CHAR_LSL_LVCC = 9999

    '***********For DC table method , we need import table method to program then we can open it***********
    If LeakPinDic.Count = 0 Then
        'GenLeakPinDic  ''' after user apply DC table method then this function could be enable
    End If
    If ContiPinDic.Count = 0 Then
        'GenContiPinDic  ''' after user apply DC table method then this function could be enable
    End If




    '***********For ADG Trace method, we need import relay trace method to program then we can open it***********
    'MatrixLoc MatrixShtName     ''' after user apply Relay trace method then we could enable this function
    'GetNameAndNum               ''' after user apply Relay trace method then we could enable this function


    '***********For ADG Trace method, we need import relay trace method to program then we can open it***********
    'g_ADG_Ctrl_DO_Flag = True   ''' after user apply Relay trace method then we could enable this function

    '***********For Limit sheet, if user have imported limit sheet relative bas/cls then we can open it***********
    Call SetCurrentLimitSet

    '    Call spotcal_Pre_OnProgramStarted

    Call Nwire_Flag_Judgement
    
    Call PrintEEPROMInfo
    
    Exit Function
ErrHandler:
    HandleExecIPError "OnProgramStarted"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


' Immediately before "post-job reset" when the test program completes.
' Note that any actions taken here with respect to modification of binning
' will affect the binning sent to the Operator Interface, but will not affect
' the binning reported in Datalog.
Function OnProgramEnded()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnProgramEnded"

    ' Put code here


    Exit Function
ErrHandler:
    HandleExecIPError "OnProgramEnded"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' Immediately before a site is disconnected.
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
Function OnPreShutDownSite()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnPreShutDownSite"

    ''    ' Put code here
    ''    If power_dcvs_exit = True Or power_dcvi_exit = True Then
    ''        PowerDown_Parallel AllPowerPinlist, AllDCVIPinlist, All_DigitalPinlist_Disc, , True ' powerdown once bin out
    ''    End If

    If InStr(UCase(TheExec.CurrentJob), "DIB") = 0 Then

        Dim PinName As String

        TheHdw.Wait 2 * ms

        PinName = "ALL_POWER"

        TheHdw.DCVI.Pins(PinName).BleederResistor = tlDCVIBleederResistorOff
        TheHdw.DCVI.Pins(PinName).Alarm = tlAlarmDefault
        TheHdw.DCVI.Pins(PinName).FoldCurrentLimit.TimeOut = 0.005


        With TheHdw.DCVI(PinName)
            .SetCurrentAndRange 1000 * mA, 1000 * mA
            TheHdw.Wait 2 * ms
            .Voltage = 0
            TheHdw.Wait 2 * ms
            .BleederResistor = tlDCVIBleederResistorOn
            TheHdw.Wait 2 * ms
            .BleederResistor = tlDCVIBleederResistorOff
            .Gate = False
        End With

    End If




    Exit Function
ErrHandler:
    HandleExecIPError "OnPreShutDownSite"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
' Immediately after a site is disconnected.
Function OnPostShutDownSite()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnPostShutDownSite"

    If InStr(UCase(TheExec.CurrentJob), "DIB") = 0 Then
        Dim PinName As String

        TheHdw.Wait 2 * ms

        PinName = "ALL_POWER"
        TheHdw.DCVI.Pins(PinName).BleederResistor = tlDCVIBleederResistorOff
        TheHdw.DCVI.Pins(PinName).Alarm = tlAlarmDefault
        TheHdw.DCVI.Pins(PinName).FoldCurrentLimit.TimeOut = 0.005


        With TheHdw.DCVI(PinName)
            .SetCurrentAndRange 1000 * mA, 1000 * mA
            TheHdw.Wait 2 * ms
            .Voltage = 0
            TheHdw.Wait 2 * ms
            .BleederResistor = tlDCVIBleederResistorOn
            TheHdw.Wait 2 * ms
            .BleederResistor = tlDCVIBleederResistorOff
            .Gate = False
        End With

    End If

    Exit Function
ErrHandler:
    HandleExecIPError "OnPostShutDownSite"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' Immediately befoe any new calibration factors are loaded
' or new calibrations run.  Not called if no action is taken during AutoCal.
Function OnAutoCalStarted()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnAutoCalStarted"

    ' Put code here


    Exit Function
ErrHandler:
    HandleExecIPError "OnAutoCalStarted"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' Immediately after AutoCal has completed.
' Not called no action has been taken (new factors loaded, or cal performed).
Function OnAutoCalCompleted()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnAutoCalCompleted"

    ' Put code here


    Exit Function
ErrHandler:
    HandleExecIPError "OnAutoCalCompleted"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


' Called right before an alarm is reported
' The alarmList is a tab delimited string of alarm error messages
Function OnAlarmOccurred(alarmList As String)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnAlarmOccurred"

    '    UNCOMMENT TO THE FOLLOWING LINES TO PARSE ALARMS

    Dim i         As Long
    Dim alarmArray() As String
    Dim s         As Long
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
ErrHandler:
    HandleExecIPError "OnAlarmOccurred"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' When the user pressed the VB Stop button, this interpose function would be called after OnPostShutDownSite was called.
' The user would put code here to make sure global variable are created and contain the correct data.
Function OnGlobalVariableReset()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnGlobalVariableReset"

    ' Put code here

    Exit Function
ErrHandler:
    HandleExecIPError "OnGlobalVariableReset"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' When the user pressed the VB Stop button, this interpose function would be called after OnPostShutDownSite was called.
' The user would put code here to make sure global variable are created and contain the correct data.
Function OnDSPGlobalVariableReset()    'otp_template 20190321
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnDSPGlobalVariableReset"
    'If THEEXEC.Sites.Active.Count = 0 Then Exit Function 'to prevent to clear again during job running 20190522
    Call ResetDspGlobalVariable

    Exit Function
ErrHandler:
    HandleExecIPError "OnDSPGlobalVariableReset"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function


' Immediately once Vaildation get started
Function OnValidationStart()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnValidationStart"


    Application.ScreenUpdating = False
    ''    '***********For Limit sheet, if user have imported limit sheet relative bas/cls then we can open it***********
    ActivateAllSheet
    Call LoadCurrentLimitSet

    'Must set to true to avoid excel screen freeze when validation fails
    Application.ScreenUpdating = True

    Exit Function
ErrHandler:
    HandleExecIPError "OnValidationStart"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
' Immediately at the conclusion of the workbook close process. The function is called in any of the following options,
' File->Close
' File->Exit
' Directly triggered the close (?X?) button of the workbook.
Function OnProgramClose()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "OnProgramClose"

    ' Put code here
    TheHdw.DIB.PowerOn = False  'reset DIB power

    Exit Function
ErrHandler:
    HandleExecIPError "OnProgramClose"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Init_Datalog_Setup()    'in on program validated
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Init_Datalog_Setup"

    TheExec.Datalog.Setup.DatalogSetup.PartResult = True
    TheExec.Datalog.Setup.DatalogSetup.XYCoordinates = True
    TheExec.Datalog.Setup.DatalogSetup.DisableChannelNumberInPTR = True    'disable channel name to stdf, PE's datalog request -- 131225, chihome
    TheExec.Datalog.Setup.DatalogSetup.OutputWidth = 0
    If LCase(TheExec.CurrentJob) Like "*char*" Then
        TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.TestName.Width = 75
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.TestName.Width = 60
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 70
    End If
    TheExec.Datalog.ApplySetup  'must need to apply after datalog setup

    Exit Function
ErrHandler:
    'TheExec.AddOutput "Error:: Init_Datalog_Setup"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Ignore_nWire_Error()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Ignore_nWire_Error"
    TheExec.Error.Behavior("HSDMPI:1335") = tlErrorIgnore
    TheExec.Error.Behavior("HSDMPI:0109") = tlErrorIgnore
    TheExec.Error.Behavior("NWirePI:0074") = tlErrorIgnore
    Exit Function
ErrHandler:
    'TheExec.AddOutput "Error:: Ignore_nWire_Error"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsModuleExist(ModuleName As String) As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IsModuleExist"
    IsModuleExist = False
    Dim mdl       As Object
    For Each mdl In Application.ThisWorkbook.VBProject.VBComponents
        If (mdl.Name = ModuleName) Then
            IsModuleExist = True
            Exit For
        End If
    Next mdl

    Exit Function
ErrHandler:
    'TheExec.AddOutput "Error:: IsModuleExist = " + ModuleName
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Init_DIB_Power()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Init_DIB_Power"
    TheHdw.DIB.PowerOn = False  'reset DIB power
    TheHdw.Wait 3# * ms
    TheHdw.DIB.Power.Item("12V").State = tlOn
    TheHdw.DIB.Power.Item("5V_1").State = tlOn
    TheHdw.DIB.Power.Item("5V_2").State = tlOn
    TheHdw.DIB.Power.Item("3.3V").State = tlOn
    TheHdw.Wait 5# * ms
    Exit Function
ErrHandler:
    'TheExec.AddOutput "Error:: Init_DIB_Power"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
