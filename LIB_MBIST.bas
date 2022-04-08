Attribute VB_Name = "LIB_MBIST"
Option Explicit
Public gS_currPayload_pattSetName As String
Public MatchFlag As Boolean
''''-----------------------------------------------------------------
Public gB_findCpuMbist_flag As Boolean
Public gB_findGpuMbist_flag As Boolean
Public gB_findSocMbist_flag As Boolean
Public gB_findPwrPin_flag As Boolean
Public gS_SocMbist_sheetName As String
Public gS_CpuMbist_sheetName As String
Public gS_GpuMbist_sheetName As String
Public gl_burst_pat As New Dictionary
Public Flag_BurstPat_INIT As Boolean ''carter 20191118

'=======================20160301=======================================
Public Function auto_FuncTest_Mbist_ExecuteForShowFailBlock(m_pattname As String, EnableBinOut As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_FuncTest_Mbist_ExecuteForShowFailBlock"
    
    ''''-------------------------------------------------------------------------------------------------
    ''''20151020 Update (Check for Mbist Function)
    ''''-------------------------------------------------------------------------------------------------
    Dim m_tn As Long
    Dim m_tn_restore As Long
    Dim m_tn_BurstIndex As Long
    Dim site As Variant
    ''''-------------------------------------------------------------------------------------------------
    ''''20160301 Update (Check for MBISTFailBlock)
    ''''-------------------------------------------------------------------------------------------------
    Dim numcap As Long
    Dim numPrecap As Long
    Dim Mbist_repair_vector As Long
    Dim Mbist_repair_cycle As Long
    Dim k As Long, Count As Long, j As Long
    
    Dim TestPatName As String, rtnPatternNames() As String, rtnPatternCount As Long
    Dim PatternNamesArray() As String
    
    Dim mem_location As String, i As Long
    Dim instanceName As String
    Dim maxDepth As Integer
'    Dim Shift_Pat As Pattern
    Dim patt As Variant
    Dim PatternName As String
    Dim PassOrFail As New SiteLong
    Dim MBISTFailBlockFlag As Boolean
    Dim PMAndBlock As String
    Dim allpins As String
    Dim blJump As Boolean
    Dim Pins As New PinData
    Dim PinData As New PinListData
    Dim Cdata As Variant
    Dim Temp As Long
    Dim m_testName As String
'    Dim IsLargeThanMaxDepth As Boolean
    Dim FailCount As New PinListData
    Dim blPatPass As New SiteBoolean
    Dim lFlagIdx As Long
    Dim astrPattPathSplit() As String
    Dim strPattName As String
    Dim blMbistFP_Binout As Boolean
    Dim lGetFlagIdx As Long
    
    blMbistFP_Binout = EnableBinOut And gl_MbistFP_Binout       '' 20160629  webster
    m_tn_BurstIndex = 2                                        '' for Mbist finger print test number(for pattern burst case, interval between payload)
    
    ''''-------------------------------------------------------------------------------------------------
    ''''20151102, Reset F_Payload Every time before runing payload
    Dim m_flagname As String
    Count = 0
    m_flagname = "F_Payload"
    allpins = "JTAG_TDO"
'    Shift_Pat = m_pattname
    For Each site In TheExec.sites.Existing
        TheExec.sites.Item(site).FlagState(m_flagname) = logicFalse ''''mean Pass
    Next site
    gS_currPayload_pattSetName = m_pattname ''''for SONE datalog
    ''''-------------------------------------------------------------------------------------------------
    m_testName = TheExec.DataManager.instanceName
    instanceName = LCase(m_testName)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    maxDepth = TheHdw.Digital.HRAM.maxDepth
    TheHdw.Digital.HRAM.Size = maxDepth
    TheHdw.Digital.HRAM.CaptureType = captFail
    ''PMAndBlock = Mid(m_testName, InStr(m_testName, "_MC") + 1, 9)
''    GetPatListFromPatternSet m_pattname, rtnPatternNames, rtnPatternCount
    Call PATT_GetPatListFromPatternSet(m_pattname, rtnPatternNames, rtnPatternCount)
    
    blPatPass = True '' 20160629
For Each patt In rtnPatternNames
    TheHdw.Digital.HRAM.SetTrigger trigFail, False, 0, True
    TheHdw.Digital.Patgen.ClearFail
    
    astrPattPathSplit = Split(CStr(patt), "\")
    strPattName = UCase(astrPattPathSplit(UBound(astrPattPathSplit)))
    If strPattName Like "*.GZ" Then strPattName = Replace(strPattName, ".GZ", "")
    
    MatchFlag = False
    For Temp = 0 To UBound(tpCycleBlockInfor)
        If UCase(tpCycleBlockInfor(Temp).strPattName) = strPattName Then
            tpEvaPattCycleBlockInfor = tpCycleBlockInfor(Temp).tpMbistCycleBlock
            MatchFlag = True
            Exit For
        End If
    Next Temp
    
    TheHdw.Patterns(patt).start
    TheHdw.Digital.Patgen.HaltWait


    ''blPatPass = TheHdw.Digital.Patgen.PatternBurstPassed
    numcap = TheHdw.Digital.HRAM.CapturedCycles

'    FailCount = TheHdw.Digital.Pins(AllPins).FailCount         '' webster add 20160428

    For Each site In TheExec.sites
        
        m_tn = TheExec.sites.Item(site).TestNumber
        m_tn_restore = m_tn

        If TheHdw.Digital.Patgen.PatternBurstPassed(site) = True Then
            Call TheExec.Datalog.WriteFunctionalResult(site, m_tn, logTestPass, , m_testName)
            
            If blPatPass(site) <> False Then  '''20160624 for pattern group, PrePattern not fail
                TheExec.sites.Item(site).testResult = sitePass  ''''20160506
            End If
            
            blPatPass(site) = True
             
'            If (UCase(m_BinFlagName) <> UCase("Default")) Then
'                If (TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicTrue) Then
'                    ''''<Important>
'                    ''''Because it was Failed on previous test, so it will NOT do any change here.
'                Else
'                    TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicFalse ''''mean Pass
'                End If
'            End If
        Else
            ''''Fail/Alarm Case
            Call TheExec.Datalog.WriteFunctionalResult(site, m_tn, logTestFail, , m_testName)
            
            TheExec.sites.Item(site).testResult = siteFail ''''20151112 update
            blPatPass(site) = False
'            If (UCase(m_BinFlagName) <> UCase("Default")) Then
'                TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicTrue ''''mean Fail
'            End If
'            ''''20151102, for SONE
'            TheExec.Sites.Item(Site).FlagState(m_flagname) = logicTrue ''''mean Fail
        End If

        TheExec.sites.Item(site).TestNumber = m_tn * 1000 + 1
            
    Next site
    
    If MatchFlag = False Then
        TheExec.Datalog.WriteComment ("Warning!! Pattern Name not match ")
        Exit For
    End If
        
    If MatchFlag And blMbistFP_Binout Then
        For k = 0 To UBound(tpEvaPattCycleBlockInfor)
            If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                lGetFlagIdx = GetFlagInfoArrIndex(tpEvaPattCycleBlockInfor(k).strFlagName)
                If lGetFlagIdx >= 0 Then
                    tyFlagInfoArr(lGetFlagIdx).CheckInfo = True
                End If
            End If
        Next k
    End If
        
    For Each site In TheExec.sites
        If blPatPass(site) = False Then     ''  patt fail
        
        
            For i = 0 To UBound(tpEvaPattCycleBlockInfor)
                tpEvaPattCycleBlockInfor(i).strFlagName = ""
            Next i
            
            For i = 0 To numcap - 1
          ''  For i = 0 To FailCount(Site) - 1 ''  webster add 20160428
                Set PinData = TheHdw.Digital.Pins(allpins).HRAM.PinData(i)
                Mbist_repair_vector = TheHdw.Digital.HRAM.PatGenInfo(i, pgVector)
                Mbist_repair_cycle = TheHdw.Digital.HRAM.PatGenInfo(i, pgCycle)
                'Mbist_repair_vector = Mbist_repair_vector + 1 'no shift
'                mem_location = "Not Match"
                'Array selection
                For Each Pins In PinData.Pins
                    Cdata = Pins.Value(site)
                    If instanceName Like "*bist*" Then
                        For j = 0 To UBound(tpEvaPattCycleBlockInfor)
                            If Mbist_Repair_CompareType = "Cycle" Then
                                If Mbist_repair_cycle = tpEvaPattCycleBlockInfor(j).lCycle Then
                                    MBISTFailBlockFlag = True
                                    Exit For
                                End If
                            ElseIf Mbist_Repair_CompareType = "Vector" Then
                                If Mbist_repair_vector = tpEvaPattCycleBlockInfor(j).lVector Then
                                    MBISTFailBlockFlag = True
                                    Exit For
                                End If
                            End If
                        Next j
                        If MBISTFailBlockFlag Then
                            MBISTFailBlockFlag = False
                            If Mbist_Repair_CompareType = "Cycle" Then
                                 For k = Count To UBound(tpEvaPattCycleBlockInfor)
                                    If Mbist_repair_vector = tpEvaPattCycleBlockInfor(k).lCycle Then
                                        If tpEvaPattCycleBlockInfor(k).strCompare <> Cdata Then
                                            PassOrFail(site) = 0
                                            tpEvaPattCycleBlockInfor(k).strFlagName = "fail"
                                            If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                 TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                                            End If
                                        Else
                                            PassOrFail(site) = 1
                                            tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
                                            If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                     TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                End If
                                            End If
                                        End If
                                        blJump = True
                                        Count = k + 1
                                    Else
                                        PassOrFail(site) = 1
                                        tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                            If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                 TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                            End If
                                        End If
                                    End If
                                TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                        tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                    If blJump = True Then
                                        blJump = False
                                        Exit For
                                    End If
                                Next k
                            ElseIf Mbist_Repair_CompareType = "Vector" Then
                                For k = Count To UBound(tpEvaPattCycleBlockInfor)
                                    If Mbist_repair_vector = tpEvaPattCycleBlockInfor(k).lVector Then
                                        If tpEvaPattCycleBlockInfor(k).strCompare <> Cdata Then
                                            PassOrFail(site) = 0
                                            tpEvaPattCycleBlockInfor(k).strFlagName = "fail"
                                            If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                 TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                                            End If
                                        Else
                                            PassOrFail(site) = 1
                                            tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
                                            If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                     TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                End If
                                            End If
                                        End If
                                        blJump = True
                                        Count = k + 1
                                    Else
                                        PassOrFail(site) = 1
                                        tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                            If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                 TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                            End If
                                        End If
                                    End If
                                    TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                                                tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(i).lVector, , , , , tlForceNone
                                    If blJump = True Then
                                        blJump = False
                                        Exit For
                                    End If
                                Next k
                            End If
                        End If
                    End If
                Next Pins
            Next i
            
                       
            
'''            ''' ===========================
'''            If numcap = maxDepth Then  '' HRAM is full
'''                TheHdw.Digital.Patgen.HaltMode = tlHaltOnHRAMFull
'''                TheHdw.Digital.HRAM.Size = maxDepth
'''                TheHdw.Digital.HRAM.CaptureType = captFail
'''                TheHdw.Digital.HRAM.SetTrigger trigFail, False, 0, True
'''
'''' '''               TheHdw.Digital.Patgen.Events.SetCycleCount True, Mbist_repair_vector + 1 ' from last fail cycle+1
'''                TheHdw.Digital.Patgen.MaskTilVector = True
'''
'''                For Each patt In rtnPatternNames
'''                    TheHdw.Patterns(m_pattname).start
'''                    TheHdw.Digital.Patgen.HaltWait
'''                Next patt
'''                Dim numcap_1 As Long
'''                numcap_1 = TheHdw.Digital.HRAM.CapturedCycles
'''                If numcap_1 Then
'''                    IsLargeThanMaxDepth = True
'''                Else
'''                    IsLargeThanMaxDepth = False
'''                End If
'''                TheHdw.Digital.Patgen.MaskTilVector = False
'''            End If
'''
'''            If k < UBound(tpEvaPattCycleBlockInfor) Then
'''                If IsLargeThanMaxDepth = False Then
'''                    PassOrFail(Site) = 1
'''                    For k = Count To UBound(tpEvaPattCycleBlockInfor)
'''                        Theexec.Flow.TestLimit PassOrFail(Site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
'''                                tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
'''                    Next k
'''                Else
'''                    Theexec.Datalog.WriteComment ("Warning!! The pattern fail cycle exceed HRAM maxDepth: " & maxDepth)
'''                End If       ' If IsLargeThanMaxDepth
'''            End If
'''
'''            If IsLargeThanMaxDepth = False Then
'''                Theexec.Flow.TestLimit 1, 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
'''                                           "Pattern_fail_cycle_size_check", , , , , tlForceNone
'''            Else
'''                Theexec.Flow.TestLimit 0, 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
'''                                           "Pattern_fail_cycle_size_check", , , , , tlForceNone
'''            End If
            
            If k < UBound(tpEvaPattCycleBlockInfor) Then        '' in unread all info of  tpEvaPattCycleBlockInfor case
                If numcap < maxDepth Then
'                    PassOrFail(Site) = 1
                    For k = Count To UBound(tpEvaPattCycleBlockInfor)
                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                            If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                 TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                            End If
                        End If
'                        TheExec.Flow.TestLimit PassOrFail(Site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
'                                tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
                    Next k
                Else
                    '' add for HRAM is full and still have some cycles need to judge, to set all flag status = true
                    If gl_MbistFP_Binout Then
                        For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                            If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                        Next k
                    End If
                End If
            End If
            
            If numcap >= maxDepth Then   '' HRAM is full
                TheExec.Flow.TestLimit 0, , , , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                          "Fail_cycle_size_check", , , , , tlForceNone
                TheExec.Datalog.WriteComment ("The number of pattern fail cycles full or exceed HRAM maxDepth: " & maxDepth)
            Else
                TheExec.Flow.TestLimit 1, , , , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                          "Fail_cycle_size_check", , , , , tlForceNone
            End If

            TheExec.sites.Item(site).TestNumber = m_tn_restore + m_tn_BurstIndex
            Count = 0
            k = 0
'            IsLargeThanMaxDepth = False
'        End If

        Else    ''blPatPass(Site) = True
            If blMbistFP_Binout Then
                For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                    If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                        If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                            TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                        End If
                    End If
                Next k
            End If
            TheExec.sites.Item(site).TestNumber = m_tn_restore + m_tn_BurstIndex
        End If '' If blPatPass(Site)
    Next site
Next patt

   '' TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode 'recover halt on opcode
'    TheExec.Flow.IncrementTestNumber
    ''''-------------------------------------------------------------------------------------------------
    
    Dim m_instName As String
    Dim FlowTestName() As String
    m_instName = TheExec.DataManager.instanceName
    If UCase(m_instName) Like UCase("*RING*") Then
        Dim MeasF_Pin As New PinList
        MeasF_Pin.Value = "RINGS_RO_CLK_OUT"
        'Call HardIP_FrequencyMeasure(MeasureF_Pin_SingleEnd, False, TestLimitPerPin_VFI, LowLimitVal(0), HighLimitVal(0), TestSeqNum, Pat, Flag_SingleLimit, d_MeasF_Interval, MeasF_WaitTime, MeasF_EventSource)
        'Call HardIP_FrequencyMeasure(MeasF_Pin, False, "FFF", 0, 0, 0, m_pattname, True, 0.01, FlowTestName)
    End If
    auto_FuncTest_Mbist_ExecuteForShowFailBlock = 1
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function MbistRampApplyLevel_AutoReadingContext(Optional ByVal ApplyPins As String = "CorePower", Optional RampingStep As Double = 10, Optional RampWaitTime As Double = 0, Optional instanceName As String)

    ''SWLINZA20171120, for ramping voltage for each ATPG and Mbist instance

    Dim Apply_Pins_Ary() As String
    Dim Apply_Pins_count As Long
    Dim Extra_RampingTime As Double: Extra_RampingTime = RampWaitTime 'RampDown_Time = 0
    'Dim RampingStep As Double
    Dim Original_voltage() As Double
    Dim Apply_TargetVoltage() As Double
    Dim DiffVoltage() As Double
    Dim RampingVoltage() As Double
    Dim Voltage_from_HW As String
    Dim i, j As Integer
    Dim Current_DCCategory As String
    Dim Current_DCSelector As String
    Dim TestBlock As String
    Dim SepcSymbolic As String
    Dim ApplyPins_String As String
    Dim ApplyPins_Boolean() As Boolean
    'Dim AllPins_needApply As Boolean
    Dim Dummy_tempStr As String
    
    If TheExec.EnableWord("Ramping_MbistATPG") = False Then Exit Function
    
    TheExec.DataManager.DecomposePinList ApplyPins, Apply_Pins_Ary(), Apply_Pins_count
    ReDim Original_voltage(Apply_Pins_count - 1) As Double
    ReDim DiffVoltage(Apply_Pins_count - 1) As Double
    ReDim RampingVoltage(Apply_Pins_count - 1) As Double
    ReDim Apply_TargetVoltage(Apply_Pins_count - 1) As Double
    ReDim ApplyPins_Boolean(Apply_Pins_count - 1) As Boolean
    
    '----- to get target voltage from DC spec for each instance -----
    'Apply_TargetVoltage
    
    'Swlinza 20180126, to save test time in IGXL9.0, use this command instead of following two
    TheExec.DataManager.GetInstanceContext Current_DCCategory, Current_DCSelector, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr, Dummy_tempStr
    'Current_DCCategory = TheExec.TestInstances.Item(InstanceName).TimingAndLevels.DCCategory
    'Current_DCSelector = TheExec.TestInstances.Item(InstanceName).TimingAndLevels.DCSelector
    
    If Current_DCCategory = Previous_DCCategory And Current_DCSelector = Previous_DCSelector Then
        Exit Function
    Else
        Previous_DCCategory = Current_DCCategory
        Previous_DCSelector = Current_DCSelector
    End If
    
    TestBlock = Mid(Current_DCCategory, 1, 3)
    
    Select Case UCase(TestBlock)
        Case UCase("Soc")
            SepcSymbolic = "_VAR_S"
        Case UCase("Cpu")
            SepcSymbolic = "_VAR_C"
        Case UCase("Gfx")
            SepcSymbolic = "_VAR_G"
        Case UCase("RTO")
            SepcSymbolic = "_VAR_R"
        Case Else
            SepcSymbolic = "_VAR_H"
    End Select
    
    '------ to calculate ramping voltage for each pins ------
    'AllPins_needApply = False
    For i = 0 To Apply_Pins_count - 1
        Original_voltage(i) = FormatNumber(TheHdw.DCVS.Pins(Apply_Pins_Ary(i)).Voltage.Main, 3)
        Apply_TargetVoltage(i) = TheExec.specs.DC.Item(Apply_Pins_Ary(i) & SepcSymbolic).Categories.Item(Current_DCCategory).Selectors.Item(Current_DCSelector).ContextValue
        DiffVoltage(i) = Original_voltage(i) - Apply_TargetVoltage(i)
        RampingVoltage(i) = FormatNumber((DiffVoltage(i) / RampingStep), 3)
        If Apply_TargetVoltage(i) = Original_voltage(i) Or Abs(DiffVoltage(i)) < 0.001 * RampingStep Then
            ApplyPins_Boolean(i) = False
        Else
            ApplyPins_Boolean(i) = True
            'AllPins_needApply = True
        End If
    Next i
    'If AllPins_needApply = False Then Exit Function
    '--------- Ramp down for retention voltage ------'
    For i = 0 To RampingStep - 1
        For j = 0 To Apply_Pins_count - 1
            If ApplyPins_Boolean(j) = True Then
                If i = RampingStep - 1 Then
                    TheHdw.DCVS.Pins(Apply_Pins_Ary(j)).Voltage.Main = Apply_TargetVoltage(j)
                Else
                    TheHdw.DCVS.Pins(Apply_Pins_Ary(j)).Voltage.Main = Original_voltage(j) - RampingVoltage(j) * i
                End If
            End If
        Next j
        TheHdw.Wait Extra_RampingTime / RampingStep
    Next i
    
End Function

 Public Function Finger_print(pattern_load As String, RunFailCycle As Boolean, Optional Flag_Name As String, Optional mbist_loop As Boolean = False)
 
    Dim maxDepth As Integer
    Dim patt As Variant
    Dim site As Variant
 
    Dim rtnPatternNames() As String, rtnPatternCount As Long
    Dim astrPattPathSplit() As String
    Dim astrPattPathSplit_01() As String
    Dim blPatPass As New SiteBoolean
    Dim numcap As Long
    Dim PinData_d As New PinListData
    Dim Mbist_repair_cycle As Long
    Dim Pins As New PinData
    Dim Cdata As Variant
    Dim TestNumber As New SiteLong
    Dim ins_new_name As String
    Dim tested As New SiteBoolean
    Dim strPattName As String
    Dim inst_match As Boolean
    Dim Temp As Long
    Dim allpins As String
    Dim PinData As New PinListData

    Dim LogLen As Long
    Dim LogLimited As Long
    Dim PrintTimes As Integer
    Dim PrintIdx As Integer
    Dim DecomposeLog() As String
    Dim ReviseStr As String

    Dim blMbistFP_Binout As Boolean
    Dim MBISTFailBlockFlag As Boolean
    Dim PassOrFail As New SiteLong
    Dim lGetFlagIdx As Long
    Dim blJump As Boolean
    Dim m_testName As String
    Dim k As Long, p As Long, g As Long, j As Long, i As Long:: k = 0:: p = 0:: g = 0:: j = 0:: i = 0

    Dim m_tn As Long
    Dim m_tn_restore As Long
    Dim m_tn_BurstIndex As Long
    Dim Mbist_repair_vector As Long
    
    '----------------------------------------------------------------
    ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08
    '----------------------------------------------------------------
    Dim Pattern_Desc As String
    Dim Pattern_Server As String
    Dim CurCount_FailAry_Element As New SiteLong
    Dim NewFmt_Printing_Header As String
    Dim CharacterNumbers As Long
    Dim Pattern_GenericName() As String
    Dim MFP_pattern_idx As Long
    Dim MFP_flow_idx As Variant
    
    m_tn_BurstIndex = 2
    allpins = "JTAG_TDO"
    LogLimited = 255
    m_testName = TheExec.DataManager.instanceName
    Call PATT_GetPatListFromPatternSet(pattern_load, rtnPatternNames, rtnPatternCount)
        
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    maxDepth = TheHdw.Digital.HRAM.maxDepth
    TheHdw.Digital.HRAM.Size = maxDepth
    TheHdw.Digital.HRAM.CaptureType = captFail
    
    TheHdw.Digital.HRAM.SetTrigger trigFail, False, 0, True
    TheHdw.Digital.Patgen.ClearFail
        
        For Each patt In rtnPatternNames
            '==================================================================='''''''''''finger print_block_01_begin
            If TheExec.EnableWord("Mbist_FingerPrint") = True Then
                astrPattPathSplit = Split(CStr(patt), "\")
                strPattName = UCase(astrPattPathSplit(UBound(astrPattPathSplit)))
    '            If strPattName Like "*:*" Then
    '                astrPattPathSplit_01 = Split(strPattName, ":")
    '                strPattName = astrPattPathSplit_01(0)
    '            End If
                If strPattName Like "*.GZ" Then strPattName = Replace(strPattName, ".GZ", "")
                
                '---------------------------------------------
                'SWLINZA 20181128 for MFP DTR, to split patset
                '---------------------------------------------
                Pattern_GenericName() = Split(strPattName, ":")
                Pattern_GenericName(0) = UCase(Pattern_GenericName(0))
                If Pattern_GenericName(0) Like "*.PAT" Then Pattern_GenericName(0) = Replace(Pattern_GenericName(0), ".PAT", "")
                    
                MatchFlag = False
                For Temp = 0 To UBound(tpCycleBlockInfor)
                    If UCase(tpCycleBlockInfor(Temp).strPattName) = strPattName Then
                        tpEvaPattCycleBlockInfor = tpCycleBlockInfor(Temp).tpMbistCycleBlock
                        MatchFlag = True
                        MFP_pattern_idx = Temp 'SWLINZa 20180907 for MFP DTR, to indentify pattern#
                        Exit For
                    End If
                Next Temp
            End If
            '==================================================================='''''''''''finger print_block_01_end
            Call TheHdw.Patterns(patt).Test(pfAlways, 0, tlResultModeDomain)
            numcap = TheHdw.Digital.HRAM.CapturedCycles
            '//////////////////////////////////////////////////////////////////////////////////////////////////'''''''''''finger print_block_03_begin
            If RunFailCycle = True And TheExec.EnableWord("Mbist_FingerPrint") = True Then

                If MatchFlag = False Then
                    TheExec.Datalog.WriteComment ("Warning!! Pattern Name not match ")
                    'Exit Function
                'End If
                Else
                    If MatchFlag And blMbistFP_Binout Then
                        For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                            If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                lGetFlagIdx = GetFlagInfoArrIndex(tpEvaPattCycleBlockInfor(k).strFlagName)
                                If lGetFlagIdx >= 0 Then
                                    tyFlagInfoArr(lGetFlagIdx).CheckInfo = True
                                End If
                            End If
                        Next k
                    End If
                    '========================================================================================
                    For Each site In TheExec.sites
                        m_tn = TheExec.sites.Item(site).TestNumber
                        m_tn_restore = m_tn
                        TheExec.sites.Item(site).TestNumber = m_tn * 1000 + 1
                        '############################################################################################
                        blPatPass(site) = TheHdw.Digital.Patgen.PatternBurstPassed
                        If blPatPass(site) = False Then     ''  patt fail
                        
                            '--------------------------------------------------------------------------
                            'SWLINZA 20181128 for MFP DTR, to get "flow-condition"(form flow) and compose DTR header
                            '--------------------------------------------------------------------------
                            MFP_flow_idx = TheExec.sites.Item(site).SiteVariableValue("MFP_Flow_Idx")
                            Pattern_Desc = tpCycleBlockInfor(MFP_pattern_idx).strDecsName(MFP_flow_idx)
                            Pattern_Server = tpCycleBlockInfor(MFP_pattern_idx).strServerName(MFP_flow_idx)
                            NewFmt_Printing_Header = "MemFP,1" & "," & site & "," & Pattern_Server & "," & Pattern_Desc & "," & UCase(m_testName) & "," & Pattern_GenericName(0) & ","
                            CharacterNumbers = LogLimited - Len(NewFmt_Printing_Header)
                            If CharacterNumbers <= 0 Then
                                TheExec.Datalog.WriteComment "The length of header is over than" & LogLimited & "."
                                TheExec.Datalog.WriteComment "Please  Check the length of header which consist of instance name and pattern name."
                            End If
                            Dim Pattern_Failure_Cycles() As New SiteVariant
                            For i = 0 To UBound(Pattern_Failure_Cycles())
                                Pattern_Failure_Cycles(i) = ""
                            Next i
                            CurCount_FailAry_Element = 0

                            For i = 0 To numcap - 1
                                Set PinData = TheHdw.Digital.Pins(allpins).HRAM.PinData(i)
                                Mbist_repair_cycle = TheHdw.Digital.HRAM.PatGenInfo(i, pgCycle)
                                   For Each Pins In PinData.Pins
                                        Cdata = Pins.Value(site)
                                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                        If TheExec.DataManager.instanceName Like "*bist*" Then
                                            For j = 0 To UBound(tpEvaPattCycleBlockInfor)
                                                If Mbist_Repair_CompareType = "Cycle" Then
                                                    Mbist_repair_cycle = TheHdw.Digital.HRAM.PatGenInfo(i, pgCycle)
                                                    If Mbist_repair_cycle = tpEvaPattCycleBlockInfor(j).lCycle Then
                                                        MBISTFailBlockFlag = True
                                                        Exit For
                                                    End If
                                                ElseIf Mbist_Repair_CompareType = "Vector" Then
                                                    Mbist_repair_vector = TheHdw.Digital.HRAM.PatGenInfo(i, pgVector)
                                                    If Mbist_repair_vector = tpEvaPattCycleBlockInfor(j).lVector Then
                                                        MBISTFailBlockFlag = True
                                                        Exit For
                                                    End If
                                                End If
                                            Next j
                                            If MBISTFailBlockFlag Then
                                                    MBISTFailBlockFlag = False
                                                   '=================================================================================
                                                    If Mbist_Repair_CompareType = "Cycle" Then
                                                    'If Mbist_repair_cycle = tpEvaPattCycleBlockInfor(k).lCycle Then
                                                            For k = Count To UBound(tpEvaPattCycleBlockInfor)
                                                                If Mbist_repair_cycle = tpEvaPattCycleBlockInfor(k).lCycle Then
                                                                    If tpEvaPattCycleBlockInfor(k).strCompare <> Cdata Then
                                                                        PassOrFail(site) = 0
                                                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                                             TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                                                                        End If
                                                                                                                    
                                                                        '---------------------------------------------------------------------------------------------
                                                                        ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08/M,
                                                                        ' to count Fail Cycles and store in ary for print later
                                                                        ' Array element + Header must lower than 255, if it's over than 255 then print warning message
                                                                        '----------------------------------------------------------------------------------------------
                                                                        If Len("," & CStr(Mbist_repair_cycle)) < CharacterNumbers Then
                                                                            If Len(Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(Mbist_repair_cycle)) > CharacterNumbers Then
                                                                                CurCount_FailAry_Element = CurCount_FailAry_Element + 1
                                                                            Else
                                                                                CurCount_FailAry_Element = CurCount_FailAry_Element
                                                                            End If
                                                                            If CurCount_FailAry_Element > 50 Then
                                                                                ReDim Preserve Pattern_Failure_Cycles(CurCount_FailAry_Element)
                                                                            End If
                                                                            If Pattern_Failure_Cycles(CurCount_FailAry_Element) = "" Then
                                                                                Pattern_Failure_Cycles(CurCount_FailAry_Element) = CStr(Mbist_repair_cycle)
                                                                            Else
                                                                                Pattern_Failure_Cycles(CurCount_FailAry_Element) = Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(Mbist_repair_cycle)
                                                                            End If
                                                                        Else
                                                                            TheExec.Datalog.WriteComment "The Remianing Character Numbers is not enough to output failing cycles."
                                                                            TheExec.Datalog.WriteComment "Please check the length of header which consist of instance name and pattern name."
                                                                        End If
                                                                        
                                                                    Else
                                                                        PassOrFail(site) = 1
                                                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                                            If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                                                 TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                                            End If
                                                                        End If
                                                                    End If
                                                                    blJump = True
                                                                    Count = k + 1
                                                                Else
                                                                    PassOrFail(site) = 1
                                                                    If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                                        If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                                             TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                                        End If
                                                                    End If
                                                                End If
                                                                
                                                                ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                                                                LogLen = Len(ReviseStr)
                                                                If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                                                                For PrintIdx = 0 To PrintTimes - 1
                                                                    ReDim Preserve DecomposeLog(PrintTimes - 1)
                                                                    DecomposeLog(PrintIdx) = Mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                                                                    TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                                                Next PrintIdx
        '''                                                                TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
        '''                                                                        tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                                                If blJump = True Then
                                                                    blJump = False
                                                                    Exit For
                                                                End If
                                                            Next k
                                            '=================================================================================
                                            ElseIf Mbist_Repair_CompareType = "Vector" Then
                                                        For k = Count To UBound(tpEvaPattCycleBlockInfor)
                                                            If Mbist_repair_vector = tpEvaPattCycleBlockInfor(k).lVector Then
                                                                If tpEvaPattCycleBlockInfor(k).strCompare <> Cdata Then
                                                                    PassOrFail(site) = 0
                                                                    tpEvaPattCycleBlockInfor(k).strFlagName = "fail"
                                                                    If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                                         TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                                                                    End If
                                                                                                                
                                                                    '---------------------------------------------------------------------------------------------
                                                                    ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08/M
                                                                    ' to count Fail vectors and store in ary for print later
                                                                    ' Array element + Header must lower than 255, if it's over than 255 then print warning message
                                                                    '----------------------------------------------------------------------------------------------
                                                                    If Len("," & CStr(Mbist_repair_vector)) < CharacterNumbers Then
                                                                        If Len(Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(Mbist_repair_vector)) > CharacterNumbers Then
                                                                            CurCount_FailAry_Element = CurCount_FailAry_Element + 1
                                                                        Else
                                                                            CurCount_FailAry_Element = CurCount_FailAry_Element
                                                                        End If
                                                                        If CurCount_FailAry_Element > 50 Then
                                                                            ReDim Preserve Pattern_Failure_Cycles(CurCount_FailAry_Element)
                                                                        End If
                                                                        If Pattern_Failure_Cycles(CurCount_FailAry_Element) = "" Then
                                                                            Pattern_Failure_Cycles(CurCount_FailAry_Element) = CStr(Mbist_repair_vector)
                                                                        Else
                                                                            Pattern_Failure_Cycles(CurCount_FailAry_Element) = Pattern_Failure_Cycles(CurCount_FailAry_Element) & "," & CStr(Mbist_repair_vector)
                                                                        End If
                                                                    Else
                                                                        TheExec.Datalog.WriteComment "The Remianing Character Numbers is not enough to output failing vectors."
                                                                        TheExec.Datalog.WriteComment "Please check the length of header which consist of instance name and pattern name."
                                                                    End If

                                                                Else
                                                                    PassOrFail(site) = 1
                                                                    tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
                                                                    If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                                        If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                                             TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                                        End If
                                                                    End If
                                                                End If
                                                                blJump = True
                                                                Count = k + 1
                                                            Else
                                                                PassOrFail(site) = 1
                                                                tpEvaPattCycleBlockInfor(k).strFlagName = "pass"
                                                                If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                                    If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                                         TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                                    End If
                                                                End If
                                                            End If
                                                            
                                                            
                                                                ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                                                                LogLen = Len(ReviseStr)
                                                                If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                                                                For PrintIdx = 0 To PrintTimes - 1
                                                                    ReDim Preserve DecomposeLog(PrintTimes - 1)
                                                                    DecomposeLog(PrintIdx) = Mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                                                                    TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
                                                                Next PrintIdx
        '''                                                                TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
        '''                                                                        tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lVector, , , , , tlForceNone
                                                                If blJump = True Then
                                                                    blJump = False
                                                                    Exit For
                                                                End If
                                                        Next k
                                            End If
                                            '=================================================================================
                                            End If
                                        End If
                                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                   Next Pins
                            Next i
                            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                            If k < UBound(tpEvaPattCycleBlockInfor) Then        '' in unread all info of  tpEvaPattCycleBlockInfor case
                                If numcap < maxDepth Then
                                    PassOrFail(site) = 1
                                    For k = Count To UBound(tpEvaPattCycleBlockInfor)
                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                            If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                 TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                            End If
                                        End If
                                        
                                                ReduceBlkLen tpEvaPattCycleBlockInfor(k).strBlaclName, ReviseStr
                                                LogLen = Len(ReviseStr)
                                                If LogLen Mod LogLimited <> 0 Then PrintTimes = (LogLen \ LogLimited) + 1
                                                For PrintIdx = 0 To PrintTimes - 1
                                                    ReDim Preserve DecomposeLog(PrintTimes - 1)
                                                    DecomposeLog(PrintIdx) = Mid(ReviseStr, 1 + PrintIdx * LogLimited, LogLimited)
                                                    TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                            DecomposeLog(PrintIdx) & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                                Next PrintIdx
                                
'''                                                TheExec.Flow.TestLimit PassOrFail(site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_testName, , _
'''                                                        tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                    Next k
                                Else
                                    '' add for HRAM is full and still have some cycles need to judge, to set all flag status = true
                                    If gl_MbistFP_Binout Then
                                        For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                                            If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                                        Next k
                                    End If
                                End If
                            End If
                            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                            If numcap >= maxDepth Then   '' HRAM is full
                                TheExec.Flow.TestLimit 0, , , , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                          "Fail_cycle_size_check", , , , , tlForceNone
                                TheExec.Datalog.WriteComment ("The number of pattern fail cycles full or exceed HRAM maxDepth: " & maxDepth)
                            Else
                                TheExec.Flow.TestLimit 1, , , , , scaleNoScaling, , , "MemFP_" & m_testName, , _
                                                          "Fail_cycle_size_check", , , , , tlForceNone
                            End If
                            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                            TheExec.sites.Item(site).TestNumber = m_tn_restore + m_tn_BurstIndex
                            Count = 0
                            k = 0
                            
                            '----------------------------------------------------------------
                            ' SWLINZA 20181128, for MemFP DTR, C651/Si requests 2018/08/M
                            ' To print final DTR,
                            ' Pattern_Failure_Cycles(AryIdx) is stored in previous procedure
                            '----------------------------------------------------------------
                            Dim AryIdx As Long
                            TheExec.Datalog.WriteComment ""
                            For AryIdx = 0 To CurCount_FailAry_Element
                                    TheExec.Datalog.WriteComment NewFmt_Printing_Header & Pattern_Failure_Cycles(AryIdx)
                            Next AryIdx
                            TheExec.Datalog.WriteComment ""
                            
                        Else    ''blPatPass(Site) = True
                            If blMbistFP_Binout Then
                                For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                                    If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                        If TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                            TheExec.sites.Item(site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                        End If
                                    End If
                                Next k
                            End If
                        End If '' If blPatPass(Site)
                        '############################################################################################
                    Next site
                End If
                '========================================================================================
            End If
            '//////////////////////////////////////////////////////////////////////////////////////////////'''''''''''finger print_block_03_end
            If mbist_loop Then
                '===================================================================
                For Each site In TheExec.sites
                    'testnumber(Site) = TheExec.sites.Item(Site).testnumber
                    tested(site) = False
                    blPatPass(site) = TheHdw.Digital.Patgen.PatternBurstPassed
                    '-------------------------------------------------------------------------------------------------
                    If blPatPass(site) = False Or alarmFail(site) = True Then   'pattern test fail or alarm
                        TheExec.sites.Item(site).FlagState(Flag_Name) = logicTrue 'pattern test fail
                        TheExec.sites.Item(site).testResult = siteFail
                        tested(site) = True
                        'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
                   '-------------------------------------------------------------------------------------------------
                    Else    'blPatPass(Site) = True ; pattern test pass
                        If (tested(site) = False) Then
                            If (TheExec.sites.Item(site).FlagState(Flag_Name) <> logicTrue) Then 'confirm flag is true(pattern fail)
                                TheExec.sites.Item(site).FlagState(Flag_Name) = logicFalse       'pattern test pass
                            End If
                            TheExec.sites.Item(site).testResult = sitePass
                        End If
                            'Call TheExec.Datalog.WriteFunctionalResult(Site, testnumber(Site), logTestPass, , ins_new_name)
                            'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
                    End If  '' If blPatPass(Site) End
                    '-------------------------------------------------------------------------------------------------
                    'TheExec.Datalog.WriteComment "Instance                = " & ins_new_name
                    'TheExec.Datalog.WriteComment "Pat Name                = " & m_pattname
                    'TheExec.Datalog.WriteComment "Test Falg               =>" & flag_name & "(" & Site & ") = " & TheExec.sites.Item(Site).FlagState(flag_name) & ",     if pattern pass=> flag is logicFalse => 0" & ",     if pattern fail=> flag is logicTrue => 1"
                    blPatPass(site) = False
                    alarmFail(site) = False
                Next site
                '===================================================================
            End If
        Next patt
        
Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ReduceBlkLen(InpStr As String, OutpStr As String)
    'Dim InputStr As String: InputStr = "Proc48_Mem121_Mem120_Mem123"   '_Mem119_Mem118"
    Dim SplitStr() As String
    Dim BlkFirstStr As String
    Dim BlkSecondStr As String
    Dim ReplaceStr As String
    Dim ReduceStr As String
    Dim BlkIdx As Long
    Dim DupIdx As Long
    Dim DupLen As Long
    Dim DupStr As String
    Dim IdxStr As String, i As Integer
    
    SplitStr = Split(InpStr, "_")
    If UBound(SplitStr) > 1 Then
    
    BlkFirstStr = SplitStr(0)
    BlkSecondStr = SplitStr(1)
    ReplaceStr = Right(InpStr, Len(InpStr) - Len(BlkFirstStr) - Len(BlkSecondStr) - 2)
        For DupIdx = 0 To Len(BlkSecondStr) - 1
        IdxStr = Mid(BlkSecondStr, DupIdx + 1, 1)
            If IsNumeric(IdxStr) Then
                DupLen = DupIdx
                DupStr = Left(BlkSecondStr, DupLen)
                Exit For
            End If

        Next
    ReduceStr = Replace(ReplaceStr, DupStr, "")
    OutpStr = BlkFirstStr & "_" & BlkSecondStr & "_" & ReduceStr
    

    Else
        OutpStr = InpStr

    End If
End Function

Public Function Parsing_Busrt_Pattern()
''''Start, modify from T-sic, Carter, 20191106
    On Error GoTo errHandler
    
    Dim burst_pat() As String
    
    If Flag_BurstPat_INIT = False Then
        Dim i As Long
        Dim j As Long
        Dim maxcol As Long
        Dim MaxRow As Long
        Dim sheet_idx As Long
        Dim burst_idx As Long: burst_idx = 1
        Dim start_col As Integer: start_col = 2
        Dim start_row As Integer: start_row = 3
        
        Dim arr_content() As Variant
        Dim sheetnames() As String
        
        Dim Pat_sheet As Worksheet
        
        ReDim burst_pat(burst_idx)
        
        For Each Pat_sheet In Worksheets
            If Pat_sheet.Name Like "Patsets_*" Then ''Patsets_CpuScan/Patsets_GfxScan/Patsets_SocScan //Patsets_*Scan
                Worksheets(Pat_sheet.Name).Activate
                'Debug.Print Pat_sheet.name
                MaxRow = Worksheets(Pat_sheet.Name).UsedRange.Rows.Count
                maxcol = Worksheets(Pat_sheet.Name).UsedRange.Columns.Count
                arr_content = Worksheets(Pat_sheet.Name).range(Cells(1, 1), Cells(MaxRow, maxcol)).Value
                For i = start_row To MaxRow - 1
                    If arr_content(i, 2) <> burst_pat(burst_idx - 1) And arr_content(i, 7) Like LCase("yes") Then
                        ReDim Preserve burst_pat(burst_idx)
                        burst_pat(burst_idx) = arr_content(i, 2)
                        burst_idx = burst_idx + 1
                    ElseIf arr_content(i, 2) = "" Then
                        Exit For
                    End If
                Next i
            End If
        Next Pat_sheet
        
        For i = 0 To UBound(burst_pat)
            If burst_pat(i) <> "" Then
                If Not gl_burst_pat.Exists(burst_pat(i)) Then
                    gl_burst_pat.Add burst_pat(i), UCase(burst_pat(i))
                End If
            End If
        Next i
    End If
    
    Flag_BurstPat_INIT = True
    
Exit Function
errHandler:
        TheExec.AddOutput "Error in the VBT Parsing_Busrt_Pattern"
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ATPG_offline(pattern_load As String, ResultMode As tlResultMode)
    ''Carter, 20191120
    On Error GoTo errHandler
    
    Dim Pins As Variant
    Dim patt As Variant
    
    Dim PinName() As String
    Dim m_testName As String
    Dim rtnPatternNames() As String
    
    Dim offline_patallpass As Boolean
    Dim offline_pat_status As New SiteBoolean
    
    Dim NumberPins As Long
    Dim rtnPatternCount As Long
    
    Dim Core_Vmain As Double

    m_testName = TheExec.DataManager.instanceName
    offline_patallpass = True
    offline_pat_status = False
    TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
    Call TheExec.DataManager.DecomposePinList("CorePower", PinName(), NumberPins)
    If InStr(pattern_load, "\") = 0 Then ''Burst Pattern, exclude walkingZ pattern
        Call PATT_GetPatListFromPatternSet(pattern_load, rtnPatternNames, rtnPatternCount)
        For Each patt In rtnPatternNames
            Call ATPG_offline_Simulation(patt, ResultMode, offline_patallpass, offline_pat_status)
        Next patt
    
    Else ''Single Pattern
        Call ATPG_offline_Simulation(pattern_load, ResultMode, offline_patallpass, offline_pat_status)
    
    End If
    
Exit Function
errHandler:
        TheExec.AddOutput "Error in the VBT ATPG_offline"
        If AbortTest Then Exit Function Else Resume Next
End Function


Public Function ATPG_offline_Simulation(pattern_load As Variant, ResultMode As tlResultMode, offline_patallpass As Boolean, offline_pat_status As SiteBoolean)
        
    
    If LCase(pattern_load) Like "*_in*" Then
        Call TheHdw.Patterns(pattern_load).Test(pfAlways, 0, ResultMode)
    Else
        offline_patallpass = True
              
        If EnableWord_Golden_Default = False Then
            For Each site In TheExec.sites
                If offline_pat_status(site) = False Then
                    offline_pat_status(site) = IIf(Round(WorksheetFunction.Min(1, Rnd * 30), 0) = 1, True, False)
                End If
            Next site
        Else
                offline_pat_status = True
        End If
        
        For Each site In TheExec.sites
            offline_patallpass = offline_patallpass And offline_pat_status(site)
        Next site
        
        If offline_patallpass = True Then
            Call TheHdw.Patterns(pattern_load).Test(pfAlways, 0, ResultMode)
    
        Else
            Call TheHdw.Patterns(pattern_load).Test(pfNever, 0, ResultMode)
            For Each site In TheExec.sites
                If offline_pat_status(site) = False Then
                    Call TheExec.Datalog.WriteFunctionalResult(site, TheExec.sites.Item(site).TestNumber, logTestFail)
    
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(site, TheExec.sites.Item(site).TestNumber, logTestPass)
    
                End If
            Next
        End If
    End If
    
Exit Function
errHandler:
        TheExec.AddOutput "Error in the VBT ATPG_offline_pat"
        If AbortTest Then Exit Function Else Resume Next

End Function
