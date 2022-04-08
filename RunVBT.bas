Attribute VB_Name = "RunVBT"
' This ALWAYS GENERATED file contains wrappers for VBT tests.
' Do not edit.

Public Sub COD_Check()

End Sub

Private Sub HandleUntrappedError()
    ' Sanity clause
    If TheExec Is Nothing Then
        MsgBox "IG-XL is not running!  VBT tests cannot execute unless IG-XL is running."
        Exit Sub
    End If
    ' If the last site has failed out, let's ignore the error
    If TheExec.sites.Active.Count = 0 Then Exit Sub  ' don't log the error
    ' If in a legacy site loop, make sure to complete it. (For-Each site syntax in IG-XL 6.10 aborts gracefully.)
    Do While TheExec.sites.InSiteLoop
        Call TheExec.sites.SelectNext(loopTop) '  Legacy syntax (hidden)
    Loop
    ' Select all active sites in case a subset of sites was selected when error occurred.
    TheExec.sites.Selected = TheExec.sites.Active
    ' Log the error to the IG-XL Error logging mechanism (tells Flow to fail the test)
    AbortTest
End Sub

Public Function DCVIPowerSupply_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New InterposeName
    p7.Value = v(6)
    Dim p8 As New Pattern
    p8.Value = v(7)
    Dim p9 As New PinList
    p9.Value = v(8)
    Dim p10 As New PinList
    p10.Value = v(9)
    Dim p11 As New PinList
    p11.Value = v(10)
    Dim p12 As New PinList
    p12.Value = v(11)
    Dim p13 As New PinList
    p13.Value = v(17)
    Dim p14 As New PinList
    p14.Value = v(18)
    Dim p15 As tlPSSource
    p15 = v(19)
    Dim p16 As tlRelayMode
    p16 = v(34)
    Dim p17 As New PinList
    p17.Value = v(35)
    Dim p18 As New PinList
    p18.Value = v(36)
    Dim p19 As tlPSTestControl
    p19 = v(37)
    Dim p20 As New InterposeName
    p20.Value = v(39)
    Dim p21 As tlWaitVal
    p21 = v(41)
    Dim p22 As tlWaitVal
    p22 = v(42)
    Dim p23 As tlWaitVal
    p23 = v(43)
    Dim p24 As tlWaitVal
    p24 = v(44)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    DCVIPowerSupply_T__ = Template.VBT_DCVIPowerSupply_T.DCVIPowerSupply_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, CDbl(v(12)), CLng(v(13)), CStr(v(14)), CDbl(v(15)), CDbl(v(16)), p13, p14, p15, CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CStr(v(32)), CBool(v(33)), p16, p17, p18, p19, CBool(v(38)), p20, CStr(v(40)), p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(46)), , CStr(v(47)), CBool(v(48)), CBool(v(49)), pStep)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function DCVSPowerSupply_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New InterposeName
    p7.Value = v(6)
    Dim p8 As New Pattern
    p8.Value = v(7)
    Dim p9 As New PinList
    p9.Value = v(8)
    Dim p10 As New PinList
    p10.Value = v(9)
    Dim p11 As New PinList
    p11.Value = v(10)
    Dim p12 As New PinList
    p12.Value = v(11)
    Dim p13 As New PinList
    p13.Value = v(12)
    Dim p14 As New PinList
    p14.Value = v(16)
    Dim p15 As tlPSSource
    p15 = v(17)
    Dim p16 As tlRelayMode
    p16 = v(31)
    Dim p17 As New PinList
    p17.Value = v(32)
    Dim p18 As New PinList
    p18.Value = v(33)
    Dim p19 As tlPSTestControl
    p19 = v(34)
    Dim p20 As tlWaitVal
    p20 = v(35)
    Dim p21 As tlWaitVal
    p21 = v(36)
    Dim p22 As tlWaitVal
    p22 = v(37)
    Dim p23 As tlWaitVal
    p23 = v(38)
    Dim p24 As New FormulaArg
    p24.Value = v(40)
    Dim p25 As New FormulaArg
    p25.Value = v(41)
    Dim p26 As New FormulaArg
    p26.Value = v(42)
    Dim p27 As New FormulaArg
    p27.Value = v(43)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    DCVSPowerSupply_T__ = Template.VBT_DCVSPowerSupply_T.DCVSPowerSupply_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CDbl(v(13)), CLng(v(14)), CStr(v(15)), p14, p15, CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CBool(v(30)), p16, p17, p18, p19, p20, p21, p22, p23, CBool(v(UBound(v))), p24, p25, p26, p27, , CStr(v(44)), CBool(v(45)), CBool(v(46)), pStep)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Empty_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New InterposeName
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New PinList
    p7.Value = v(12)
    Dim p8 As New PinList
    p8.Value = v(13)
    Dim p9 As New PinList
    p9.Value = v(14)
    Dim p10 As New PinList
    p10.Value = v(15)
    Dim p11 As New PinList
    p11.Value = v(16)
    Dim p12 As New PinList
    p12.Value = v(17)
    Dim p13 As New PinList
    p13.Value = v(18)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Empty_T__ = Template.VBT_Empty_T.Empty_T(p1, p2, p3, p4, p5, p6, CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CStr(v(11)), p7, p8, p9, p10, p11, p12, p13, pStep, CBool(v(19)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Functional_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New InterposeName
    p7.Value = v(6)
    Dim p8 As PFType
    p8 = v(7)
    Dim p9 As tlResultMode
    p9 = v(8)
    Dim p10 As New PinList
    p10.Value = v(9)
    Dim p11 As New PinList
    p11.Value = v(10)
    Dim p12 As New PinList
    p12.Value = v(11)
    Dim p13 As New PinList
    p13.Value = v(12)
    Dim p14 As New PinList
    p14.Value = v(13)
    Dim p15 As New PinList
    p15.Value = v(20)
    Dim p16 As New PinList
    p16.Value = v(21)
    Dim p17 As New InterposeName
    p17.Value = v(22)
    Dim p18 As tlRelayMode
    p18 = v(24)
    Dim p19 As tlWaitVal
    p19 = v(27)
    Dim p20 As tlWaitVal
    p20 = v(28)
    Dim p21 As tlWaitVal
    p21 = v(29)
    Dim p22 As tlWaitVal
    p22 = v(30)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim p23 As tlPatConcurrentMode
    p23 = v(34)
    Dim p24 As tlTemplateScanFailDataLogging
    p24 = v(35)
    Dim p25 As tlDigitalCMEMCaptureLimitMode
    p25 = v(36)
    Dim p26 As tlTemplateScanPinListSource
    p26 = v(38)
    Dim p27 As New PinList
    p27.Value = v(39)
    Dim p28 As tlTemplateScanCaptureFormat
    p28 = v(40)
    Dim p29 As tlTemplateScanCaptureDataType
    p29 = v(41)
    Dim p30 As tlTemplateScanUserCommentSource
    p30 = v(42)
    Dim p31 As tlTemplateATPGPinMapSource
    p31 = v(44)
    Functional_T__ = Template.VBT_Functional_T.Functional_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), p15, p16, p17, CStr(v(23)), p18, CBool(v(25)), CBool(v(26)), p19, p20, p21, p22, CBool(v(UBound(v))), CStr(v(32)), pStep, CStr(v(33)), p23, p24, p25, CLng(v(37)), p26, p27, p28, p29, p30, CStr(v(43)), p31, CStr(v(45)), CStr(v(46)), CBool(v(47)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function PinPMU_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New InterposeName
    p1.Value = v(1)
    Dim p2 As New InterposeName
    p2.Value = v(2)
    Dim p3 As New InterposeName
    p3.Value = v(3)
    Dim p4 As New InterposeName
    p4.Value = v(4)
    Dim p5 As New InterposeName
    p5.Value = v(5)
    Dim p6 As New InterposeName
    p6.Value = v(6)
    Dim p7 As New Pattern
    p7.Value = v(7)
    Dim p8 As New Pattern
    p8.Value = v(8)
    Dim p9 As New PinList
    p9.Value = v(10)
    Dim p10 As New PinList
    p10.Value = v(11)
    Dim p11 As New PinList
    p11.Value = v(12)
    Dim p12 As New PinList
    p12.Value = v(13)
    Dim p13 As New PinList
    p13.Value = v(14)
    Dim p14 As New PinList
    p14.Value = v(15)
    Dim p15 As tlPPMUMode
    p15 = v(16)
    Dim p16 As New FormulaArg
    p16.Value = v(18)
    Dim p17 As New FormulaArg
    p17.Value = v(19)
    Dim p18 As tlPPMURelayMode
    p18 = v(20)
    Dim p19 As New PinList
    p19.Value = v(36)
    Dim p20 As New PinList
    p20.Value = v(37)
    Dim p21 As tlWaitVal
    p21 = v(38)
    Dim p22 As tlWaitVal
    p22 = v(39)
    Dim p23 As tlWaitVal
    p23 = v(40)
    Dim p24 As tlWaitVal
    p24 = v(41)
    Dim p25 As tlPPMUMode
    p25 = v(49)
    Dim p26 As New FormulaArg
    p26.Value = v(52)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim p27 As New PinList
    p27.Value = v(53)
    Dim p28 As tlPPMUMode
    p28 = v(54)
    Dim p29 As New FormulaArg
    p29.Value = v(55)
    PinPMU_T__ = Template.VBT_PinPmu_T.PinPMU_T(CStr(v(0)), p1, p2, p3, p4, p5, p6, p7, p8, CStr(v(9)), p9, p10, p11, p12, p13, p14, p15, CDbl(v(17)), p16, p17, p18, CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CLng(v(32)), CBool(v(33)), CStr(v(34)), CStr(v(35)), p19, p20, p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(43)), CStr(v(44)), , CStr(v(45)), CBool(v(46)), CBool(v(47)), CBool(v(48)), p25, CStr(v(50)), CStr(v(51)), p26, pStep, p27, p28, p29, CStr(v(56)), CStr(v(57)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function MtoMemory_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New InterposeName
    p7.Value = v(6)
    Dim p8 As PFType
    p8 = v(7)
    Dim p9 As New PinList
    p9.Value = v(8)
    Dim p10 As New PinList
    p10.Value = v(9)
    Dim p11 As New PinList
    p11.Value = v(10)
    Dim p12 As New PinList
    p12.Value = v(11)
    Dim p13 As New PinList
    p13.Value = v(12)
    Dim p14 As New PinList
    p14.Value = v(19)
    Dim p15 As New PinList
    p15.Value = v(20)
    Dim p16 As New InterposeName
    p16.Value = v(21)
    Dim p17 As tlRelayMode
    p17 = v(24)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim ExtraArgs(0 To 49) As Variant
    Dim i As Integer
    For i = 0 To 49
        ExtraArgs(i) = v(51 + i)
    Next i
    MtoMemory_T__ = Template.VBT_MTOMemory_T.MtoMemory_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), p14, p15, p16, CStr(v(22)), CBool(v(23)), p17, CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CLng(v(29)), CStr(v(30)), CStr(v(31)), CStr(v(32)), CStr(v(33)), CLng(v(34)), CStr(v(35)), CStr(v(36)), CStr(v(37)), CStr(v(38)), CLng(v(39)), CLng(v(40)), CBool(v(UBound(v))), pStep, ExtraArgs, CStr(v(42)), CStr(v(43)), CStr(v(44)), CStr(v(45)), CStr(v(46)), CStr(v(47)), CStr(v(48)), CStr(v(49)), CStr(v(50)), CBool(v(51)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Relay_Control__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Relay_Control__ = VBAProject.VBT_LIB_Common.Relay_Control(p1, p2, CDbl(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function StartSBClock__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    StartSBClock__ = VBAProject.VBT_LIB_Common.StartSBClock(CDbl(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function StopSBClock__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    StopSBClock__ = VBAProject.VBT_LIB_Common.StopSBClock()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FreeRunclk_Enable_ori__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    FreeRunclk_Enable_ori__ = VBAProject.VBT_LIB_Common.FreeRunclk_Enable_ori(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FreeRunClk_Disable__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    FreeRunClk_Disable__ = VBAProject.VBT_LIB_Common.FreeRunClk_Disable(CStr(v(0)), CBool(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Start_Profile__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Start_Profile__ = VBAProject.VBT_LIB_Common.Start_Profile(p1, CStr(v(1)), CDbl(v(2)), CLng(v(3)), CStr(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function start_profile_DCVI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    start_profile_DCVI__ = VBAProject.VBT_LIB_Common.start_profile_DCVI(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Plot_Profile__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Plot_Profile__ = VBAProject.VBT_LIB_Common.Plot_Profile(p1, CStr(v(1)), CBool(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Plot_profile_DCVI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Plot_profile_DCVI__ = VBAProject.VBT_LIB_Common.Plot_profile_DCVI(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Start_Profile_AutoResolution__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Start_Profile_AutoResolution__ = VBAProject.VBT_LIB_Common.Start_Profile_AutoResolution(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CLng(v(3)), CStr(v(4)), CDbl(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_Footer__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_Footer__ = VBAProject.VBT_LIB_Common.Print_Footer(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_Header__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_Header__ = VBAProject.VBT_LIB_Common.Print_Header(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_PgmInfo__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_PgmInfo__ = VBAProject.VBT_LIB_Common.Print_PgmInfo(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Write_DIB_EEPROM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Write_DIB_EEPROM__ = VBAProject.VBT_LIB_Common.Write_DIB_EEPROM(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_DIB_EEPROM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_DIB_EEPROM__ = VBAProject.VBT_LIB_Common.Read_DIB_EEPROM()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReadProberTemp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ReadProberTemp__ = VBAProject.VBT_LIB_Common.ReadProberTemp(CDbl(v(0)), CDbl(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetupInitialCondition__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    SetupInitialCondition__ = VBAProject.VBT_LIB_Common.SetupInitialCondition(p1, p2, CDbl(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_PPMU_Clamp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(2)
    Dim p3 As New PinList
    p3.Value = v(4)
    Dim p4 As New PinList
    p4.Value = v(6)
    Dim p5 As New PinList
    p5.Value = v(8)
    Dim p6 As New PinList
    p6.Value = v(10)
    Dim p7 As New PinList
    p7.Value = v(12)
    Dim p8 As New PinList
    p8.Value = v(14)
    Dim p9 As New PinList
    p9.Value = v(16)
    Dim p10 As New PinList
    p10.Value = v(18)
    Set_PPMU_Clamp__ = VBAProject.VBT_LIB_Common.Set_PPMU_Clamp(p1, CDbl(v(1)), p2, CDbl(v(3)), p3, CDbl(v(5)), p4, CDbl(v(7)), p5, CDbl(v(9)), p6, CDbl(v(11)), p7, CDbl(v(13)), p8, CDbl(v(15)), p9, CDbl(v(17)), p10, CDbl(v(19)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_Package_ID__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_Package_ID__ = VBAProject.VBT_LIB_Common.Read_Package_ID()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FreeRunClk_Disable_MultiPort__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    FreeRunClk_Disable_MultiPort__ = VBAProject.VBT_LIB_Common.FreeRunClk_Disable_MultiPort(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckFlag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CheckFlag__ = VBAProject.VBT_LIB_Common.CheckFlag(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Alarm_binout__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Alarm_binout__ = VBAProject.VBT_LIB_Common.Alarm_binout()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_DCVS_alarm__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(2)
    Set_DCVS_alarm__ = VBAProject.VBT_LIB_Common.Set_DCVS_alarm(p1, CDbl(v(1)), p2, CDbl(v(3)), CBool(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function KA_start__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    KA_start__ = VBAProject.VBT_LIB_Common.KA_start()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function KA_end__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    KA_end__ = VBAProject.VBT_LIB_Common.KA_end()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Disable_compare__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Disable_compare__ = VBAProject.VBT_LIB_Common.Disable_compare(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Enble_compare__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Enble_compare__ = VBAProject.VBT_LIB_Common.Enble_compare(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Plot_Profile_on_disk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Plot_Profile_on_disk__ = VBAProject.VBT_LIB_Common.Plot_Profile_on_disk(p1, CDbl(v(1)), CLng(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SiteResultCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SiteResultCheck__ = VBAProject.VBT_LIB_Common.SiteResultCheck()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FreeRunclk_Enable__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    FreeRunclk_Enable__ = VBAProject.VBT_LIB_Common.FreeRunclk_Enable(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PowerDown_Parallel__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PowerDown_Parallel__ = VBAProject.VBT_LIB_Common.PowerDown_Parallel(CStr(v(0)), CStr(v(1)), CStr(v(2)), CDbl(v(3)), CBool(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PowerUp_Parallel__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PowerUp_Parallel__ = VBAProject.VBT_LIB_Common.PowerUp_Parallel(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CDbl(v(6)), CBool(v(7)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_Power_Alarm__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(2)
    Dim p3 As New PinList
    p3.Value = v(4)
    Set_Power_Alarm__ = VBAProject.VBT_LIB_Common.Set_Power_Alarm(p1, CDbl(v(1)), p2, CDbl(v(3)), p3, CDbl(v(5)), CBool(v(6)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Search_UnExistPin__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Search_UnExistPin__ = VBAProject.VBT_LIB_Common.Search_UnExistPin()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function License_Mapping__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    License_Mapping__ = VBAProject.VBT_LIB_Common.License_Mapping()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HIP_Init_Datalog_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    HIP_Init_Datalog_Setup__ = VBAProject.VBT_LIB_Common.HIP_Init_Datalog_Setup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Common_UnitTest__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Common_UnitTest__ = VBAProject.VBT_LIB_Common.Common_UnitTest()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function VBT_IEDA_Registry__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    VBT_IEDA_Registry__ = VBAProject.VBT_LIB_Common_AP.VBT_IEDA_Registry(CStr(v(0)), CBool(v(1)), CBool(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function IDS_eFuse_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    IDS_eFuse_Write__ = VBAProject.VBT_LIB_DC_AP.IDS_eFuse_Write(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_IDS_main_current_Delta__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    DCVS_IDS_main_current_Delta__ = VBAProject.VBT_LIB_DC_AP.DCVS_IDS_main_current_Delta(p1, CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function PPMU_Continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(4)
    Dim p3 As New PinList
    p3.Value = v(8)
    PPMU_Continuity__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Continuity(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2, CBool(v(5)), CStr(v(6)), CStr(v(7)), p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UVI80_Continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(4)
    Dim p3 As New PinList
    p3.Value = v(9)
    UVI80_Continuity__ = VBAProject.VBT_LIB_DC_Conti.UVI80_Continuity(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2, CBool(v(5)), CBool(v(6)), CDbl(v(7)), CDbl(v(8)), p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function p2p_short_Power__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    Dim p5 As tlLimitForceResults
    p5 = v(11)
    p2p_short_Power__ = VBAProject.VBT_LIB_DC_Conti.p2p_short_Power(p1, p2, p3, p4, CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), CDbl(v(7)), CDbl(v(8)), CDbl(v(9)), CDbl(v(10)), p5, CDbl(v(12)), CDbl(v(13)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Pre_PowerUp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Pre_PowerUp__ = VBAProject.VBT_LIB_DC_Conti.Pre_PowerUp(CStr(v(0)), CBool(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Conti_WalkingZ__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Conti_WalkingZ__ = VBAProject.VBT_LIB_DC_Conti.Conti_WalkingZ(p1, p2, CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PowerSensePins_continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    PowerSensePins_continuity__ = VBAProject.VBT_LIB_DC_Conti.PowerSensePins_continuity(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CStr(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Continuity_PN_Disconnect_IV_Curve__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(6)
    PPMU_Continuity_PN_Disconnect_IV_Curve__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Continuity_PN_Disconnect_IV_Curve(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), p2, CBool(v(7)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Continuity_PN_Disconnect__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(4)
    PPMU_Continuity_PN_Disconnect__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Continuity_PN_Disconnect(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2, CBool(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function p2p_short_Power_FVMI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(4)
    Dim p3 As New PinList
    p3.Value = v(6)
    p2p_short_Power_FVMI__ = VBAProject.VBT_LIB_DC_Conti.p2p_short_Power_FVMI(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2, CBool(v(5)), p3, CStr(v(7)), CStr(v(8)), CStr(v(9)), CBool(v(10)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Continuity_IV_Curve__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(6)
    PPMU_Continuity_IV_Curve__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Continuity_IV_Curve(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), p2, CBool(v(7)), CBool(v(8)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_open_measurement__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(4)
    Dim p3 As New PinList
    p3.Value = v(6)
    Power_open_measurement__ = VBAProject.VBT_LIB_DC_Conti.Power_open_measurement(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2, CBool(v(5)), p3, CStr(v(7)), CStr(v(8)), CStr(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_IO_measure_R__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As tlLimitForceResults
    p3 = v(5)
    PPMU_IO_measure_R__ = VBAProject.VBT_LIB_DC_Conti.PPMU_IO_measure_R(p1, p2, CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), p3, CBool(v(6)), CBool(v(7)), CDbl(v(8)), CDbl(v(9)), CDbl(v(10)), CDbl(v(11)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GndSensePins_continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GndSensePins_continuity__ = VBAProject.VBT_LIB_DC_Conti.GndSensePins_continuity(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function p2p_short_Power_FVMI_VI_Curve__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(6)
    p2p_short_Power_FVMI_VI_Curve__ = VBAProject.VBT_LIB_DC_Conti.p2p_short_Power_FVMI_VI_Curve(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), p2, CStr(v(7)), CStr(v(8)), CStr(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RetrieveDictionaryOfDiffPairs__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RetrieveDictionaryOfDiffPairs__ = VBAProject.VBT_LIB_DC_Conti.RetrieveDictionaryOfDiffPairs()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Measure_Contact_Resistance__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(6)
    PPMU_Measure_Contact_Resistance__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Measure_Contact_Resistance(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), p2, CBool(v(7)), CBool(v(8)), CStr(v(9)), CDbl(v(10)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Measure_Contact_Resistance_Corner_Vss__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(6)
    PPMU_Measure_Contact_Resistance_Corner_Vss__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Measure_Contact_Resistance_Corner_Vss(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), p2, CDbl(v(7)), CBool(v(8)), CBool(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function p2p_short_Power_FVMI_Parallel__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As tlLimitForceResults
    p3 = v(8)
    Dim p4 As New PinList
    p4.Value = v(10)
    p2p_short_Power_FVMI_Parallel__ = VBAProject.VBT_LIB_DC_Conti.p2p_short_Power_FVMI_Parallel(p1, p2, CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), CDbl(v(7)), p3, CBool(v(9)), p4, CStr(v(11)), CStr(v(12)), CStr(v(13)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetCurrentRange__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SetCurrentRange__ = VBAProject.VBT_LIB_DC_Conti.SetCurrentRange(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CBool(v(6)), CDbl(v(7)), CDbl(v(8)), CDbl(v(9)), CStr(v(10)), CStr(v(11)), CStr(v(12)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function DC_Func_WriteFuncResult__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DC_Func_WriteFuncResult__ = VBAProject.VBT_LIB_DC_Func.DC_Func_WriteFuncResult(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_LeakCurr_Univeral_func__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(5)
    Dim p5 As tlLimitForceResults
    p5 = v(7)
    Meas_LeakCurr_Univeral_func__ = VBAProject.VBT_LIB_DC_Func.Meas_LeakCurr_Univeral_func(p1, p2, p3, CStr(v(3)), CBool(v(4)), p4, CBool(v(6)), p5, CDbl(v(8)), CDbl(v(9)), CStr(v(10)), CDbl(v(11)), CDbl(v(12)), CBool(v(13)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_VOHL_Univeral_func__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(5)
    Dim p5 As tlLimitForceResults
    p5 = v(7)
    Meas_VOHL_Univeral_func__ = VBAProject.VBT_LIB_DC_Func.Meas_VOHL_Univeral_func(p1, p2, p3, CStr(v(3)), CBool(v(4)), p4, CBool(v(6)), p5, CDbl(v(8)), CDbl(v(9)), CStr(v(10)), CDbl(v(11)), CStr(v(12)), CStr(v(13)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_VOHL_Univeral_func_Parallel__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(5)
    Dim p5 As tlLimitForceResults
    p5 = v(7)
    Meas_VOHL_Univeral_func_Parallel__ = VBAProject.VBT_LIB_DC_Func.Meas_VOHL_Univeral_func_Parallel(p1, p2, p3, CStr(v(3)), CBool(v(4)), p4, CBool(v(6)), p5, CDbl(v(8)), CDbl(v(9)), CStr(v(10)), CDbl(v(11)), CStr(v(12)), CStr(v(13)), CInt(v(14)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_VOH_MeasVI_Univeral_func_DC__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(5)
    Dim p5 As New PinList
    p5.Value = v(6)
    Dim p6 As New PinList
    p6.Value = v(7)
    Dim p7 As New PinList
    p7.Value = v(8)
    Dim p8 As tlLimitForceResults
    p8 = v(10)
    Meas_VOH_MeasVI_Univeral_func_DC__ = VBAProject.VBT_LIB_DC_Func.Meas_VOH_MeasVI_Univeral_func_DC(p1, p2, p3, CStr(v(3)), CBool(v(4)), p4, p5, p6, p7, CBool(v(9)), p8, CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CDbl(v(19)), CDbl(v(20)), CDbl(v(21)), CDbl(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CBool(v(26)), CBool(v(27)), CBool(v(28)), CDbl(v(29)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_VOH_MeasVIR_delta_Univeral_func_DC__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(5)
    Dim p5 As New PinList
    p5.Value = v(6)
    Dim p6 As New PinList
    p6.Value = v(7)
    Dim p7 As New PinList
    p7.Value = v(8)
    Dim p8 As tlLimitForceResults
    p8 = v(10)
    Meas_VOH_MeasVIR_delta_Univeral_func_DC__ = VBAProject.VBT_LIB_DC_Func.Meas_VOH_MeasVIR_delta_Univeral_func_DC(p1, p2, p3, CStr(v(3)), CBool(v(4)), p4, p5, p6, p7, CBool(v(9)), p8, CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CDbl(v(19)), CDbl(v(20)), CDbl(v(21)), CDbl(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CBool(v(26)), CBool(v(27)), CBool(v(28)), CDbl(v(29)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Interpose_TurnOff_CPUFLagA__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' Interpose_TurnOff_CPUFLagA__ = VBAProject.VBT_LIB_DC_Func.Interpose_TurnOff_CPUFLagA(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_VIHL_VOHL_Universal_Functional__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(4)
    Meas_VIHL_VOHL_Universal_Functional__ = VBAProject.VBT_LIB_DC_Func.Meas_VIHL_VOHL_Universal_Functional(p1, p2, p3, CBool(v(3)), p4, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Cal_Hysteresis__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Cal_Hysteresis__ = VBAProject.VBT_LIB_DC_Func.Cal_Hysteresis(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Judge_GPIO_Vil__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Judge_GPIO_Vil__ = VBAProject.VBT_LIB_DC_Func.Judge_GPIO_Vil(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Judge_GPIO_Vih__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Judge_GPIO_Vih__ = VBAProject.VBT_LIB_DC_Func.Judge_GPIO_Vih(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HiZ_Leakage_Power__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(4)
    Dim p3 As New PinList
    p3.Value = v(7)
    Dim p4 As New PinList
    p4.Value = v(15)
    Dim p5 As New PinList
    p5.Value = v(16)
    HiZ_Leakage_Power__ = VBAProject.VBT_LIB_DC_Func.HiZ_Leakage_Power(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2, CDbl(v(5)), CDbl(v(6)), p3, CDbl(v(8)), CDbl(v(9)), CStr(v(10)), CStr(v(11)), CBool(v(12)), CBool(v(13)), CBool(v(14)), p4, p5)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_VIR_IO_Universal_func_GPIO__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(4)
    Dim p4 As New PinList
    p4.Value = v(11)
    Dim p5 As CalculateMethodSetup
    p5 = v(14)
    Dim p6 As New PinList
    p6.Value = v(15)
    Dim p7 As InstrumentSpecialSetup
    p7 = v(21)
    Dim p8 As CalculateMethodSetup
    p8 = v(22)
    Dim p9 As Enum_RAK
    p9 = v(23)
    Dim p10 As New InterposeName
    p10.Value = v(29)
    Dim p11 As New InterposeName
    p11.Value = v(31)
    Meas_VIR_IO_Universal_func_GPIO__ = VBAProject.VBT_LIB_DC_Func.Meas_VIR_IO_Universal_func_GPIO(p1, CStr(v(1)), CBool(v(2)), p2, p3, CBool(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), p4, CLng(v(12)), CLng(v(13)), p5, p6, CLng(v(16)), CLng(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), p7, p8, p9, CStr(v(24)), CStr(v(25)), CStr(v(26)), CBool(v(27)), CStr(v(28)), p10, CStr(v(30)), p11, CStr(v(32)), CStr(v(33)), CBool(v(34)), CInt(v(35)), CBool(v(36)), CStr(v(37)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IO_HardIP_PPMU_Measure_I_TTR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' IO_HardIP_PPMU_Measure_I_TTR__ = VBAProject.VBT_LIB_DC_Func.IO_HardIP_PPMU_Measure_I_TTR(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IO_HardIP_PPMU_Measure_V_TTR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As InstrumentSpecialSetup
    p1 = v(11)
    Dim p2 As Enum_RAK
    p2 = v(12)
    ' IO_HardIP_PPMU_Measure_V_TTR__ = VBAProject.VBT_LIB_DC_Func.IO_HardIP_PPMU_Measure_V_TTR(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_VIR_IO_Universal_func_GPIO_TTR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(4)
    Dim p4 As New PinList
    p4.Value = v(11)
    Dim p5 As CalculateMethodSetup
    p5 = v(14)
    Dim p6 As New PinList
    p6.Value = v(15)
    Dim p7 As InstrumentSpecialSetup
    p7 = v(21)
    Dim p8 As Enum_RAK
    p8 = v(23)
    Dim p9 As New InterposeName
    p9.Value = v(29)
    Dim p10 As New InterposeName
    p10.Value = v(31)
    Meas_VIR_IO_Universal_func_GPIO_TTR__ = VBAProject.VBT_LIB_DC_Func.Meas_VIR_IO_Universal_func_GPIO_TTR(p1, CStr(v(1)), CBool(v(2)), p2, p3, CBool(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), p4, CLng(v(12)), CLng(v(13)), p5, p6, CLng(v(16)), CLng(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), p7, CBool(v(22)), p8, CStr(v(24)), CStr(v(25)), CStr(v(26)), CBool(v(27)), CStr(v(28)), p9, CStr(v(30)), p10, CStr(v(32)), CStr(v(33)), CBool(v(34)), CInt(v(35)), CBool(v(36)), CStr(v(37)), CInt(v(38)), CStr(v(39)), CStr(v(40)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EvaluateEachBlock__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' EvaluateEachBlock__ = VBAProject.VBT_LIB_DC_Func.EvaluateEachBlock(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetFlowSingleUseLimit_KeepEmpty__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' GetFlowSingleUseLimit_KeepEmpty__ = VBAProject.VBT_LIB_DC_Func.GetFlowSingleUseLimit_KeepEmpty(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_Static_Power_Ramp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EVS_Static_Power_Ramp__ = VBAProject.VBT_LIB_DC_Func.EVS_Static_Power_Ramp(CStr(v(0)), CDbl(v(1)), CStr(v(2)), CInt(v(3)), CDbl(v(4)), CBool(v(5)), CDbl(v(6)), CStr(v(7)), CStr(v(8)), CBool(v(9)), CBool(v(10)), CInt(v(11)), CBool(v(12)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_Pre_Setting__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' EVS_Pre_Setting__ = VBAProject.VBT_LIB_DC_Func.EVS_Pre_Setting(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Evs_Ramp_UPorDown__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' Evs_Ramp_UPorDown__ = VBAProject.VBT_LIB_DC_Func.Evs_Ramp_UPorDown(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Test_time_breakdown_End__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Test_time_breakdown_End__ = VBAProject.VBT_LIB_DC_Func.Test_time_breakdown_End(CDbl(v(0)), CBool(v(1)), CStr(v(2)), CInt(v(3)), CDbl(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Test_time_breakdown_Start__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Test_time_breakdown_Start__ = VBAProject.VBT_LIB_DC_Func.Test_time_breakdown_Start(CDbl(v(0)), CBool(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Checkboard_EVS_Probe_Location__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_Checkboard_EVS_Probe_Location__ = VBAProject.VBT_LIB_DC_Func.auto_Checkboard_EVS_Probe_Location()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function DCVS_IDD_dynamic__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New Pattern
    p3.Value = v(2)
    DCVS_IDD_dynamic__ = VBAProject.VBT_LIB_DC_IDS.DCVS_IDD_dynamic(p1, p2, p3, CDbl(v(3)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_IDS_main_auto_range_and_measure__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(7)
    Dim p2 As New PinList
    p2.Value = v(8)
    Dim p3 As New PinList
    p3.Value = v(9)
    Dim p4 As New PinList
    p4.Value = v(10)
    Dim p5 As New PinList
    p5.Value = v(11)
    Dim p6 As New PinList
    p6.Value = v(12)
    Dim p7 As New PinList
    p7.Value = v(13)
    ' DCVS_IDS_main_auto_range_and_measure__ = VBAProject.VBT_LIB_DC_IDS.DCVS_IDS_main_auto_range_and_measure(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVI_IDS_main_auto_range_and_measure__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(6)
    Dim p2 As New PinList
    p2.Value = v(7)
    Dim p3 As New PinList
    p3.Value = v(8)
    Dim p4 As New PinList
    p4.Value = v(9)
    Dim p5 As New PinList
    p5.Value = v(10)
    ' DCVI_IDS_main_auto_range_and_measure__ = VBAProject.VBT_LIB_DC_IDS.DCVI_IDS_main_auto_range_and_measure(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IDS_main_current__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    Dim p5 As New PinList
    p5.Value = v(11)
    Dim p6 As New PinList
    p6.Value = v(12)
    Dim p7 As New PinList
    p7.Value = v(13)
    Dim p8 As New PinList
    p8.Value = v(14)
    Dim p9 As New PinList
    p9.Value = v(15)
    Dim p10 As New PinList
    p10.Value = v(16)
    Dim p11 As New PinList
    p11.Value = v(17)
    Dim p12 As New PinList
    p12.Value = v(18)
    Dim p13 As New PinList
    p13.Value = v(19)
    Dim p14 As New PinList
    p14.Value = v(20)
    Dim p15 As New PinList
    p15.Value = v(21)
    Dim p16 As New PinList
    p16.Value = v(22)
    Dim p17 As New PinList
    p17.Value = v(30)
    IDS_main_current__ = VBAProject.VBT_LIB_DC_IDS.IDS_main_current(p1, p2, p3, p4, CLng(v(4)), CBool(v(5)), CStr(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), CStr(v(10)), p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, CStr(v(23)), CBool(v(24)), CBool(v(25)), CStr(v(26)), CBool(v(27)), CStr(v(28)), CStr(v(29)), p17, CLng(v(31)), CLng(v(32)), CStr(v(33)), CStr(v(34)), CStr(v(35)), CStr(v(36)), CStr(v(37)), CStr(v(38)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Functional_T_updated__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New InterposeName
    p7.Value = v(6)
    Dim p8 As PFType
    p8 = v(7)
    Dim p9 As tlResultMode
    p9 = v(8)
    Dim p10 As New PinList
    p10.Value = v(9)
    Dim p11 As New PinList
    p11.Value = v(10)
    Dim p12 As New PinList
    p12.Value = v(11)
    Dim p13 As New PinList
    p13.Value = v(12)
    Dim p14 As New PinList
    p14.Value = v(13)
    Dim p15 As New PinList
    p15.Value = v(20)
    Dim p16 As New PinList
    p16.Value = v(21)
    Dim p17 As New InterposeName
    p17.Value = v(22)
    Dim p18 As tlRelayMode
    p18 = v(24)
    Dim p19 As CusWaitVal
    p19 = v(27)
    Dim p20 As CusWaitVal
    p20 = v(28)
    Dim p21 As CusWaitVal
    p21 = v(29)
    Dim p22 As CusWaitVal
    p22 = v(30)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim p23 As tlPatConcurrentMode
    p23 = v(34)
    Functional_T_updated__ = VBAProject.VBT_LIB_Digital_Functional_T.Functional_T_updated(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), p15, p16, p17, CStr(v(23)), p18, CBool(v(25)), CBool(v(26)), p19, p20, p21, p22, CBool(v(UBound(v))), CStr(v(32)), pStep, CStr(v(33)), p23, CStr(v(35)), CBool(v(36)), CBool(v(37)), CInt(v(38)), CStr(v(39)), CStr(v(40)), CBool(v(41)), CBool(v(42)), CLng(v(43)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DatalogType__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DatalogType__ = VBAProject.VBT_LIB_Digital_Functional_T.DatalogType()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PostTest__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' Call VBAProject.VBT_LIB_Digital_Functional_T.PostTest(*One or more unsupported types in argument list or non Long/Integer return type*)
    PostTest__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function getdefaults__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' getdefaults__ = VBAProject.VBT_LIB_Digital_Functional_T.getdefaults(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function pattern_module_test__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As PFType
    p1 = v(3)
    Dim p2 As tlResultMode
    p2 = v(5)
    Dim p3 As tlPatConcurrentMode
    p3 = v(6)
    pattern_module_test__ = VBAProject.VBT_LIB_Digital_Functional_T.pattern_module_test(CStr(v(0)), CBool(v(1)), CBool(v(2)), p1, CLng(v(4)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Mbist_Block_loop_inst_match__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_Mbist_Block_loop_inst_match__ = VBAProject.VBT_LIB_Digital_Functional_T.auto_Mbist_Block_loop_inst_match(CStr(v(0)), CStr(v(1)), CLng(v(2)), CStr(v(3)), CBool(v(4)), CBool(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Mbist_Block_loop_inst_non_match__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_Mbist_Block_loop_inst_non_match__ = VBAProject.VBT_LIB_Digital_Functional_T.auto_Mbist_Block_loop_inst_non_match(CStr(v(0)), CStr(v(1)), CLng(v(2)), CStr(v(3)), CBool(v(4)), CBool(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Init_RSCR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Init_RSCR__ = VBAProject.VBT_LIB_Digital_Mbist.Init_RSCR()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Mbist_RSCR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Mbist_RSCR__ = VBAProject.VBT_LIB_Digital_Mbist.Mbist_RSCR(p1, CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TurnOnEfusePwrPins_Mbist__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    TurnOnEfusePwrPins_Mbist__ = VBAProject.VBT_LIB_Digital_Mbist.TurnOnEfusePwrPins_Mbist(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TurnOffEfusePwrPins_Mbist__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    TurnOffEfusePwrPins_Mbist__ = VBAProject.VBT_LIB_Digital_Mbist.TurnOffEfusePwrPins_Mbist(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MbistRetentionLevelWait_and_lowDown_power__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(1)
    MbistRetentionLevelWait_and_lowDown_power__ = VBAProject.VBT_LIB_Digital_Mbist.MbistRetentionLevelWait_and_lowDown_power(CDbl(v(0)), p1, CDbl(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MbistRetentionLevelWait__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(2)
    MbistRetentionLevelWait__ = VBAProject.VBT_LIB_Digital_Mbist.MbistRetentionLevelWait(CDbl(v(0)), CDbl(v(1)), p1, CDbl(v(3)), CDbl(v(4)), CBool(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Init_MBISTFailBlock__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Init_MBISTFailBlock__ = VBAProject.VBT_LIB_Digital_Mbist.Init_MBISTFailBlock()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetFlagInfoArrIndex__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GetFlagInfoArrIndex__ = VBAProject.VBT_LIB_Digital_Mbist.GetFlagInfoArrIndex(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MbistRetentionWait__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MbistRetentionWait__ = VBAProject.VBT_LIB_Digital_Mbist.MbistRetentionWait(CDbl(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function TPmode_Char_on__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    TPmode_Char_on__ = VBAProject.VBT_LIB_Digital_Shmoo.TPmode_Char_on()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TPmode_Char_off__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    TPmode_Char_off__ = VBAProject.VBT_LIB_Digital_Shmoo.TPmode_Char_off()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function freerunclk_set_XY__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' freerunclk_set_XY__ = VBAProject.VBT_LIB_Digital_Shmoo.freerunclk_set_XY(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function freerunclk_stop__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' freerunclk_stop__ = VBAProject.VBT_LIB_Digital_Shmoo.freerunclk_stop(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CharStoreResultsUntilNextRun__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CharStoreResultsUntilNextRun__ = VBAProject.VBT_LIB_Digital_Shmoo.CharStoreResultsUntilNextRun()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function setup_patgen_counter__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' setup_patgen_counter__ = VBAProject.VBT_LIB_Digital_Shmoo.setup_patgen_counter(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function run_shmoo__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    run_shmoo__ = VBAProject.VBT_LIB_Digital_Shmoo.run_shmoo(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Functional_T_char__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New InterposeName
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As PFType
    p7 = v(6)
    Dim p8 As tlResultMode
    p8 = v(7)
    Dim p9 As New PinList
    p9.Value = v(8)
    Dim p10 As New PinList
    p10.Value = v(9)
    Dim p11 As New PinList
    p11.Value = v(10)
    Dim p12 As New PinList
    p12.Value = v(11)
    Dim p13 As New PinList
    p13.Value = v(12)
    Dim p14 As New PinList
    p14.Value = v(19)
    Dim p15 As New PinList
    p15.Value = v(20)
    Dim p16 As New InterposeName
    p16.Value = v(21)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim p17 As tlPatConcurrentMode
    p17 = v(26)
    Dim p18 As New Pattern
    p18.Value = v(28)
    Dim p19 As New Pattern
    p19.Value = v(29)
    Dim p20 As New Pattern
    p20.Value = v(30)
    Dim p21 As New Pattern
    p21.Value = v(31)
    Dim p22 As New Pattern
    p22.Value = v(32)
    Dim p23 As New Pattern
    p23.Value = v(33)
    Dim p24 As New Pattern
    p24.Value = v(34)
    Dim p25 As New Pattern
    p25.Value = v(35)
    Dim p26 As New Pattern
    p26.Value = v(36)
    Dim p27 As New Pattern
    p27.Value = v(37)
    Dim p28 As New Pattern
    p28.Value = v(38)
    Dim p29 As New Pattern
    p29.Value = v(39)
    Dim p30 As New Pattern
    p30.Value = v(40)
    Dim p31 As New Pattern
    p31.Value = v(41)
    Dim p32 As New Pattern
    p32.Value = v(42)
    Functional_T_char__ = VBAProject.VBT_LIB_Digital_Shmoo.Functional_T_char(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), p14, p15, p16, CStr(v(22)), CBool(v(UBound(v))), CStr(v(24)), pStep, CStr(v(25)), p17, CStr(v(27)), p18, p19, p20, p21, p22, p23, p24, p25, p26, p27, p28, p29, p30, p31, p32, CStr(v(43)), CStr(v(44)), CStr(v(45)), CStr(v(46)), CStr(v(47)), CStr(v(48)), CStr(v(49)), CStr(v(50)), CStr(v(51)), CBool(v(52)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PrintShmooInfo__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' PrintShmooInfo__ = VBAProject.VBT_LIB_Digital_Shmoo.PrintShmooInfo(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Flow_Shmoo_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Flow_Shmoo_Setup__ = VBAProject.VBT_LIB_Digital_Shmoo.Flow_Shmoo_Setup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function NWireFRCIsEnable__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' NWireFRCIsEnable__ = VBAProject.VBT_LIB_Digital_Shmoo.NWireFRCIsEnable(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckCharErrorCount__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CheckCharErrorCount__ = VBAProject.VBT_LIB_Digital_Shmoo.CheckCharErrorCount(CStr(v(0)), CDbl(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EnableShmooAbnormalCounter__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EnableShmooAbnormalCounter__ = VBAProject.VBT_LIB_Digital_Shmoo.EnableShmooAbnormalCounter()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DisableShmooAbnormalCounter__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DisableShmooAbnormalCounter__ = VBAProject.VBT_LIB_Digital_Shmoo.DisableShmooAbnormalCounter()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Init_Datalog_Setup_Char__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Init_Datalog_Setup_Char__ = VBAProject.VBT_LIB_Digital_Shmoo.Init_Datalog_Setup_Char()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Re_PowerOn_WhileSweep__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' Re_PowerOn_WhileSweep__ = VBAProject.VBT_LIB_Digital_Shmoo.Re_PowerOn_WhileSweep(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_PowerUp_Parallel_Interpose__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DCVS_PowerUp_Parallel_Interpose__ = VBAProject.VBT_LIB_Digital_Shmoo.DCVS_PowerUp_Parallel_Interpose(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CBool(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_PowerDown_Parallel_Interpose__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DCVS_PowerDown_Parallel_Interpose__ = VBAProject.VBT_LIB_Digital_Shmoo.DCVS_PowerDown_Parallel_Interpose(CStr(v(0)), CDbl(v(1)), CBool(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PostPointInterpose_nWire_ReStore__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' PostPointInterpose_nWire_ReStore__ = VBAProject.VBT_LIB_Digital_Shmoo.PostPointInterpose_nWire_ReStore(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function StoreMaxNum__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' StoreMaxNum__ = VBAProject.VBT_LIB_Digital_Shmoo.StoreMaxNum(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function auto_ConfigBlankChk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(3)
    Dim p4 As New PinList
    p4.Value = v(4)
    Dim p5 As New PinList
    p5.Value = v(5)
    auto_ConfigBlankChk__ = VBAProject.VBT_LIB_EFUSE_Config.auto_ConfigBlankChk(p1, p2, CStr(v(2)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ConfigWrite_byCondition__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(6)
    Dim p4 As New PinList
    p4.Value = v(7)
    Dim p5 As New PinList
    p5.Value = v(8)
    auto_ConfigWrite_byCondition__ = VBAProject.VBT_LIB_EFUSE_Config.auto_ConfigWrite_byCondition(p1, p2, CStr(v(2)), CDbl(v(3)), CStr(v(4)), CStr(v(5)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ConfigRead_by_OR_2Blocks__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(4)
    Dim p4 As New PinList
    p4.Value = v(5)
    Dim p5 As New PinList
    p5.Value = v(6)
    auto_ConfigRead_by_OR_2Blocks__ = VBAProject.VBT_LIB_EFUSE_Config.auto_ConfigRead_by_OR_2Blocks(p1, p2, CStr(v(2)), CStr(v(3)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ConfigSingleDoubleBit__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    auto_ConfigSingleDoubleBit__ = VBAProject.VBT_LIB_EFUSE_Config.auto_ConfigSingleDoubleBit(p1, p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ChkAllConfigEfuseData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_ChkAllConfigEfuseData__ = VBAProject.VBT_LIB_EFUSE_Config.auto_ChkAllConfigEfuseData(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ConfigRead_Decode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_ConfigRead_Decode__ = VBAProject.VBT_LIB_EFUSE_Config.auto_ConfigRead_Decode(p1, p2, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_eFuse_IDS_BinCut_PreCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_eFuse_IDS_BinCut_PreCheck__ = VBAProject.VBT_LIB_EFUSE_Config.auto_eFuse_IDS_BinCut_PreCheck()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_eFuse_IDS_BinCut_PostCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_eFuse_IDS_BinCut_PostCheck__ = VBAProject.VBT_LIB_EFUSE_Config.auto_eFuse_IDS_BinCut_PostCheck()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ConfigBlankChk_Early_byStage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_ConfigBlankChk_Early_byStage__ = VBAProject.VBT_LIB_EFUSE_Config.auto_ConfigBlankChk_Early_byStage(p1, p2, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_mapping_fusing_BKM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_mapping_fusing_BKM__ = VBAProject.VBT_LIB_EFUSE_Config.auto_mapping_fusing_BKM(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_getting_fusing_BKM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_getting_fusing_BKM__ = VBAProject.VBT_LIB_EFUSE_Config.auto_getting_fusing_BKM(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ConfigWrite_CFG_DV__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    auto_ConfigWrite_CFG_DV__ = VBAProject.VBT_LIB_EFUSE_Config.auto_ConfigWrite_CFG_DV(p1, CStr(v(1)), CDbl(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Parsing_BKM_Info__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Parsing_BKM_Info__ = VBAProject.VBT_LIB_EFUSE_Config.Parsing_BKM_Info(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IDS_LIMIT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    IDS_LIMIT__ = VBAProject.VBT_LIB_EFUSE_Config.IDS_LIMIT(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function auto_Efuse_DAP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_Efuse_DAP__ = VBAProject.VBT_LIB_EFUSE_DAP.auto_Efuse_DAP(p1, p2, CLng(v(2)), v(3), CStr(v(4)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Efuse_JTAG__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_Efuse_JTAG__ = VBAProject.VBT_LIB_EFUSE_DAP.auto_Efuse_JTAG(p1, p2, CLng(v(2)), v(3), CStr(v(4)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function auto_eFuse_Initialize__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_eFuse_Initialize__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_eFuse_Initialize()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ECIDBlankChk_DEID__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    Dim p5 As New PinList
    p5.Value = v(4)
    auto_ECIDBlankChk_DEID__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_ECIDBlankChk_DEID(p1, p2, p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ECIDBlankChk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(3)
    Dim p4 As New PinList
    p4.Value = v(4)
    Dim p5 As New PinList
    p5.Value = v(5)
    auto_ECIDBlankChk__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_ECIDBlankChk(p1, p2, CStr(v(2)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_EcidWrite_byCondition__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(6)
    Dim p4 As New PinList
    p4.Value = v(7)
    Dim p5 As New PinList
    p5.Value = v(8)
    auto_EcidWrite_byCondition__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_EcidWrite_byCondition(p1, p2, CStr(v(2)), CDbl(v(3)), CStr(v(4)), CStr(v(5)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ECID_Read_by_OR_2Blocks__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(4)
    Dim p4 As New PinList
    p4.Value = v(5)
    Dim p5 As New PinList
    p5.Value = v(6)
    auto_ECID_Read_by_OR_2Blocks__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_ECID_Read_by_OR_2Blocks(p1, p2, CStr(v(2)), CStr(v(3)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_EcidSingleDoubleBit__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    auto_EcidSingleDoubleBit__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_EcidSingleDoubleBit(p1, p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ReadWaferData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_ReadWaferData__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_ReadWaferData()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ReadHandlerData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_ReadHandlerData__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_ReadHandlerData()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ShowECIDData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_ShowECIDData__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_ShowECIDData()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_EcidSingleDoubleBit_nonDEID__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    auto_EcidSingleDoubleBit_nonDEID__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_EcidSingleDoubleBit_nonDEID(p1, p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ECID_Info__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ECID_Info__ = VBAProject.VBT_LIB_EFUSE_ECID.ECID_Info()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_CleanRegData_New__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_CleanRegData_New__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_CleanRegData_New()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Function_Test__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    auto_Function_Test__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_Function_Test(p1, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_eFuse_ReadAllData_to_DictDSPWave__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_eFuse_ReadAllData_to_DictDSPWave__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_eFuse_ReadAllData_to_DictDSPWave(v(0), CBool(v(1)), CBool(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_eFuse_ReadData_to_DictDSPWave_byCategory__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_eFuse_ReadData_to_DictDSPWave_byCategory__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_eFuse_ReadData_to_DictDSPWave_byCategory(v(0), CStr(v(1)), CBool(v(2)), CBool(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_eFuse_Print_GetStoredCaptureData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_eFuse_Print_GetStoredCaptureData__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_eFuse_Print_GetStoredCaptureData(CStr(v(0)), CBool(v(1)), CBool(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ECIDRead_Decode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_ECIDRead_Decode__ = VBAProject.VBT_LIB_EFUSE_ECID.auto_ECIDRead_Decode(p1, p2, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EFUSE_Resistance__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    EFUSE_Resistance__ = VBAProject.VBT_LIB_EFUSE_ECID.EFUSE_Resistance(p1, CStr(v(1)), CDbl(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function auto_MONITORBlankChk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(3)
    Dim p4 As New PinList
    p4.Value = v(4)
    Dim p5 As New PinList
    p5.Value = v(5)
    auto_MONITORBlankChk__ = VBAProject.VBT_LIB_EFUSE_MONITOR.auto_MONITORBlankChk(p1, p2, CStr(v(2)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_MONITORWrite_byCondition__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(6)
    Dim p4 As New PinList
    p4.Value = v(7)
    Dim p5 As New PinList
    p5.Value = v(8)
    auto_MONITORWrite_byCondition__ = VBAProject.VBT_LIB_EFUSE_MONITOR.auto_MONITORWrite_byCondition(p1, p2, CStr(v(2)), CDbl(v(3)), CStr(v(4)), CStr(v(5)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_MONITORRead_by_OR_2Blocks__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(5)
    Dim p4 As New PinList
    p4.Value = v(6)
    Dim p5 As New PinList
    p5.Value = v(7)
    auto_MONITORRead_by_OR_2Blocks__ = VBAProject.VBT_LIB_EFUSE_MONITOR.auto_MONITORRead_by_OR_2Blocks(p1, p2, CStr(v(2)), CStr(v(3)), CStr(v(4)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_MONITORSingleDoubleBit__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    auto_MONITORSingleDoubleBit__ = VBAProject.VBT_LIB_EFUSE_MONITOR.auto_MONITORSingleDoubleBit(p1, p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ChkAllMONITOREfuseData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_ChkAllMONITOREfuseData__ = VBAProject.VBT_LIB_EFUSE_MONITOR.auto_ChkAllMONITOREfuseData()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_MONITORRead_Decode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_MONITORRead_Decode__ = VBAProject.VBT_LIB_EFUSE_MONITOR.auto_MONITORRead_Decode(p1, p2, CBool(v(UBound(v))), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function AddStoredFuseData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' AddStoredFuseData__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.AddStoredFuseData(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetStoredFuseData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' GetStoredFuseData__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.GetStoredFuseData(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_CreateConstant__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_CreateConstant__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_CreateConstant()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_fusetofile__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_fusetofile__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_fusetofile(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_dump_fuse_data__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_dump_fuse_data__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_dump_fuse_data(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_eFuse_pgm2file__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' auto_eFuse_pgm2file__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_eFuse_pgm2file(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_eFuse_SetReadValue__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' auto_eFuse_SetReadValue__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_eFuse_SetReadValue(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_filetofuse__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_filetofuse__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_filetofuse(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Gen_ProgramBitArray__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Gen_ProgramBitArray__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.Gen_ProgramBitArray(CLng(v(0)), CLng(v(1)), CLng(v(2)), CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ConfigWrite_byCondition_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_ConfigWrite_byCondition_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_ConfigWrite_byCondition_Pgm2File(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ConfigSingleDoubleBit_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_ConfigSingleDoubleBit_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_ConfigSingleDoubleBit_Pgm2File(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRP_USI_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UDRP_USI_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_UDRP_USI_Pgm2File(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRP_USO_Syntax_Chk_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UDRP_USO_Syntax_Chk_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_UDRP_USO_Syntax_Chk_Pgm2File(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRE_USI_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UDRE_USI_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_UDRE_USI_Pgm2File(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRE_USO_Syntax_Chk_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UDRE_USO_Syntax_Chk_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_UDRE_USO_Syntax_Chk_Pgm2File(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_MONITORWrite_byCondition_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_MONITORWrite_byCondition_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_MONITORWrite_byCondition_Pgm2File(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_MONITORSingleDoubleBit_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_MONITORSingleDoubleBit_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_MONITORSingleDoubleBit_Pgm2File(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_EcidSingleDoubleBit_nonDEID_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    auto_EcidSingleDoubleBit_nonDEID_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_EcidSingleDoubleBit_nonDEID_Pgm2File(p1, p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_EcidWrite_byCondition_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_EcidWrite_byCondition_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_EcidWrite_byCondition_Pgm2File(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_EcidSingleDoubleBit_Pgm2File__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_EcidSingleDoubleBit_Pgm2File__ = VBAProject.VBT_LIB_EFUSE_Pgm2File.auto_EcidSingleDoubleBit_Pgm2File(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function auto_CMP_Syntax_Chk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_CMP_Syntax_Chk__ = VBAProject.VBT_LIB_EFUSE_UDR.auto_CMP_Syntax_Chk(p1, p2, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_USI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDR_USI__ = VBAProject.VBT_LIB_EFUSE_UDR.auto_UDR_USI(p1, p2, CStr(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_USO_Syntax_Chk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UDR_USO_Syntax_Chk__ = VBAProject.VBT_LIB_EFUSE_UDR.auto_UDR_USO_Syntax_Chk(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_UFP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    auto_UDR_UFP__ = VBAProject.VBT_LIB_EFUSE_UDR.auto_UDR_UFP(p1, CStr(v(1)), CDbl(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_UFR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    auto_UDR_UFR__ = VBAProject.VBT_LIB_EFUSE_UDR.auto_UDR_UFR(p1, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_USO_BlankChk_byStage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDR_USO_BlankChk_byStage__ = VBAProject.VBT_LIB_EFUSE_UDR.auto_UDR_USO_BlankChk_byStage(p1, p2, CStr(v(2)), CBool(v(3)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_USO_Read_Decode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDR_USO_Read_Decode__ = VBAProject.VBT_LIB_EFUSE_UDR.auto_UDR_USO_Read_Decode(p1, p2, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_USO_BlankChk_Early_byStage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDR_USO_BlankChk_Early_byStage__ = VBAProject.VBT_LIB_EFUSE_UDR.auto_UDR_USO_BlankChk_Early_byStage(p1, p2, CBool(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_USI_Sim__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UDR_USI_Sim__ = VBAProject.VBT_LIB_EFUSE_UDR.auto_UDR_USI_Sim(CBool(v(0)), CBool(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_USO_COMPARE__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDR_USO_COMPARE__ = VBAProject.VBT_LIB_EFUSE_UDR.auto_UDR_USO_COMPARE(p1, p2, CStr(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function auto_CMPE_Syntax_Chk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_CMPE_Syntax_Chk__ = VBAProject.VBT_LIB_EFUSE_UDRE.auto_CMPE_Syntax_Chk(p1, p2, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRE_USI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDRE_USI__ = VBAProject.VBT_LIB_EFUSE_UDRE.auto_UDRE_USI(p1, p2, CStr(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRE_USO_Syntax_Chk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UDRE_USO_Syntax_Chk__ = VBAProject.VBT_LIB_EFUSE_UDRE.auto_UDRE_USO_Syntax_Chk(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRE_USO_BlankChk_byStage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDRE_USO_BlankChk_byStage__ = VBAProject.VBT_LIB_EFUSE_UDRE.auto_UDRE_USO_BlankChk_byStage(p1, p2, CBool(v(2)), CStr(v(3)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRE_USO_Read_Decode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDRE_USO_Read_Decode__ = VBAProject.VBT_LIB_EFUSE_UDRE.auto_UDRE_USO_Read_Decode(p1, p2, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRE_USO_BlankChk_Early_byStage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDRE_USO_BlankChk_Early_byStage__ = VBAProject.VBT_LIB_EFUSE_UDRE.auto_UDRE_USO_BlankChk_Early_byStage(p1, p2, CBool(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRE_USI_Sim__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UDRE_USI_Sim__ = VBAProject.VBT_LIB_EFUSE_UDRE.auto_UDRE_USI_Sim(CBool(v(0)), CBool(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRE_USO_COMPARE__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDRE_USO_COMPARE__ = VBAProject.VBT_LIB_EFUSE_UDRE.auto_UDRE_USO_COMPARE(p1, p2, CStr(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function auto_CMPP_Syntax_Chk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_CMPP_Syntax_Chk__ = VBAProject.VBT_LIB_EFUSE_UDRP.auto_CMPP_Syntax_Chk(p1, p2, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRP_USI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDRP_USI__ = VBAProject.VBT_LIB_EFUSE_UDRP.auto_UDRP_USI(p1, p2, CStr(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRP_USO_Syntax_Chk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UDRP_USO_Syntax_Chk__ = VBAProject.VBT_LIB_EFUSE_UDRP.auto_UDRP_USO_Syntax_Chk(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRP_USO_BlankChk_byStage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDRP_USO_BlankChk_byStage__ = VBAProject.VBT_LIB_EFUSE_UDRP.auto_UDRP_USO_BlankChk_byStage(p1, p2, CBool(v(2)), CStr(v(3)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRP_USO_Read_Decode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDRP_USO_Read_Decode__ = VBAProject.VBT_LIB_EFUSE_UDRP.auto_UDRP_USO_Read_Decode(p1, p2, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRP_USO_BlankChk_Early_byStage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDRP_USO_BlankChk_Early_byStage__ = VBAProject.VBT_LIB_EFUSE_UDRP.auto_UDRP_USO_BlankChk_Early_byStage(p1, p2, CBool(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRP_USI_Sim__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UDRP_USI_Sim__ = VBAProject.VBT_LIB_EFUSE_UDRP.auto_UDRP_USI_Sim(CBool(v(0)), CBool(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDRP_USO_COMPARE__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    auto_UDRP_USO_COMPARE__ = VBAProject.VBT_LIB_EFUSE_UDRP.auto_UDRP_USO_COMPARE(p1, p2, CStr(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function auto_UIDBlankChk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    Dim p5 As New PinList
    p5.Value = v(4)
    auto_UIDBlankChk__ = VBAProject.VBT_LIB_EFUSE_UID.auto_UIDBlankChk(p1, p2, p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UIDBlankChk_byStage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    Dim p5 As New PinList
    p5.Value = v(4)
    auto_UIDBlankChk_byStage__ = VBAProject.VBT_LIB_EFUSE_UID.auto_UIDBlankChk_byStage(p1, p2, p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UID_Read_by_OR_2Blocks__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(4)
    Dim p4 As New PinList
    p4.Value = v(5)
    Dim p5 As New PinList
    p5.Value = v(6)
    auto_UID_Read_by_OR_2Blocks__ = VBAProject.VBT_LIB_EFUSE_UID.auto_UID_Read_by_OR_2Blocks(p1, p2, CStr(v(2)), CStr(v(3)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UIDSingleDoubleBit__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    auto_UIDSingleDoubleBit__ = VBAProject.VBT_LIB_EFUSE_UID.auto_UIDSingleDoubleBit(p1, p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UIDWrite_byCondition__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(6)
    Dim p4 As New PinList
    p4.Value = v(7)
    Dim p5 As New PinList
    p5.Value = v(8)
    auto_UIDWrite_byCondition__ = VBAProject.VBT_LIB_EFUSE_UID.auto_UIDWrite_byCondition(p1, p2, CStr(v(2)), CDbl(v(3)), CStr(v(4)), CStr(v(5)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UID_Encoding_128bits__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_UID_Encoding_128bits__ = VBAProject.VBT_LIB_EFUSE_UID.auto_UID_Encoding_128bits()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Meas_Vdiff_func__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(7)
    Dim p5 As New PinList
    p5.Value = v(10)
    Meas_Vdiff_func__ = VBAProject.VBT_LIB_HardIP.Meas_Vdiff_func(p1, p2, p3, CStr(v(3)), CBool(v(4)), CStr(v(5)), CStr(v(6)), p4, CLng(v(8)), CLng(v(9)), p5, CLng(v(11)), CLng(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_VIR_IO_Universal_func__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(4)
    Dim p4 As New PinList
    p4.Value = v(11)
    Dim p5 As CalculateMethodSetup
    p5 = v(14)
    Dim p6 As New PinList
    p6.Value = v(15)
    Dim p7 As InstrumentSpecialSetup
    p7 = v(21)
    Dim p8 As CalculateMethodSetup
    p8 = v(22)
    Dim p9 As Enum_RAK
    p9 = v(23)
    Meas_VIR_IO_Universal_func__ = VBAProject.VBT_LIB_HardIP.Meas_VIR_IO_Universal_func(p1, CStr(v(1)), CBool(v(2)), p2, p3, CBool(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), p4, CLng(v(12)), CLng(v(13)), p5, p6, CLng(v(16)), CLng(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), p7, p8, p9, CStr(v(24)), CStr(v(25)), CStr(v(26)), CBool(v(27)), CStr(v(28)), CBool(v(29)), CStr(v(30)), CStr(v(31)), CStr(v(32)), CStr(v(33)), CStr(v(34)), CStr(v(35)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_FreqVoltCurr_Universal_func__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(4)
    Dim p4 As EventSourceWithTerminationMode
    p4 = v(10)
    Dim p5 As New PinList
    p5.Value = v(17)
    Dim p6 As New PinList
    p6.Value = v(20)
    Dim p7 As CalculateMethodSetup
    p7 = v(26)
    Dim p8 As InstrumentSpecialSetup
    p8 = v(27)
    Dim p9 As Enum_RAK
    p9 = v(51)
    Meas_FreqVoltCurr_Universal_func__ = VBAProject.VBT_LIB_HardIP.Meas_FreqVoltCurr_Universal_func(p1, CStr(v(1)), CBool(v(2)), p2, p3, CBool(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), p4, CBool(v(11)), CDbl(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), p5, CLng(v(18)), CLng(v(19)), p6, CLng(v(21)), CLng(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), p7, p8, CStr(v(28)), CStr(v(29)), CStr(v(30)), CBool(v(31)), CStr(v(32)), CStr(v(33)), CBool(v(34)), CBool(v(35)), CDbl(v(36)), CDbl(v(37)), CDbl(v(38)), CDbl(v(39)), CDbl(v(40)), CDbl(v(41)), CStr(v(42)), CStr(v(43)), CStr(v(44)), CStr(v(45)), CStr(v(46)), CStr(v(47)), CStr(v(48)), CStr(v(49)), CStr(v(50)), p9, CStr(v(52)), CBool(v(53)), CBool(v(54)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Opt_DdrLpBkFunc2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New Pattern
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    Dim p5 As New PinList
    p5.Value = v(4)
    Dim p6 As New PinList
    p6.Value = v(9)
    Dim p7 As CalculateMethodSetup
    p7 = v(16)
    Opt_DdrLpBkFunc2__ = VBAProject.VBT_LIB_HardIP.Opt_DdrLpBkFunc2(p1, p2, p3, p4, p5, CLng(v(5)), CStr(v(6)), CStr(v(7)), CBool(v(8)), p6, CLng(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), p7, CStr(v(17)), CLng(v(18)), CStr(v(19)), CInt(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Opt_DdrLpBkFunc3__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New Pattern
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    Dim p5 As New PinList
    p5.Value = v(4)
    Dim p6 As New PinList
    p6.Value = v(9)
    Dim p7 As CalculateMethodSetup_DSPWave
    p7 = v(16)
    Opt_DdrLpBkFunc3__ = VBAProject.VBT_LIB_HardIP.Opt_DdrLpBkFunc3(p1, p2, p3, p4, p5, CInt(v(5)), CLng(v(6)), CLng(v(7)), CBool(v(8)), p6, CLng(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), p7, CStr(v(17)), CLng(v(18)), CStr(v(19)), CLng(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CBool(v(25)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DigSrc_DigCap_Universal_func__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(4)
    DigSrc_DigCap_Universal_func__ = VBAProject.VBT_LIB_HardIP.DigSrc_DigCap_Universal_func(p1, p2, CLng(v(2)), CLng(v(3)), p3, CLng(v(5)), CLng(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TMPS_Voltage_Print__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    TMPS_Voltage_Print__ = VBAProject.VBT_LIB_HardIP.TMPS_Voltage_Print(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReMeasImpedByAveTrimCode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(5)
    ReMeasImpedByAveTrimCode__ = VBAProject.VBT_LIB_HardIP.ReMeasImpedByAveTrimCode(CStr(v(0)), CBool(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), p1, CLng(v(6)), CStr(v(7)), CLng(v(8)), CLng(v(9)), CStr(v(10)), CStr(v(11)), CBool(v(12)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TMPS_Bin2Dec__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' TMPS_Bin2Dec__ = VBAProject.VBT_LIB_HardIP.TMPS_Bin2Dec(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TMPS_Dec2Bin__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' TMPS_Dec2Bin__ = VBAProject.VBT_LIB_HardIP.TMPS_Dec2Bin(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Eye_Diagram__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Eye_Diagram__ = VBAProject.VBT_LIB_HardIP.Eye_Diagram(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Impedance_Function__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PPMU_Impedance_Function__ = VBAProject.VBT_LIB_HardIP.PPMU_Impedance_Function(CStr(v(0)), CDbl(v(1)), CStr(v(2)), CStr(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), CDbl(v(7)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Impedance__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PPMU_Impedance__ = VBAProject.VBT_LIB_HardIP.PPMU_Impedance(CStr(v(0)), CDbl(v(1)), CStr(v(2)), CStr(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), CDbl(v(7)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeDig__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(4)
    Dim p4 As New PinList
    p4.Value = v(18)
    TrimCodeDig__ = VBAProject.VBT_LIB_HardIP.TrimCodeDig(p1, CStr(v(1)), CBool(v(2)), p2, p3, CLng(v(5)), CStr(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), CDbl(v(10)), CDbl(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), p4, CLng(v(19)), CBool(v(UBound(v))), CStr(v(21)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeDig_SeaHawk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(4)
    Dim p4 As New PinList
    p4.Value = v(18)
    TrimCodeDig_SeaHawk__ = VBAProject.VBT_LIB_HardIP.TrimCodeDig_SeaHawk(p1, CStr(v(1)), CBool(v(2)), p2, p3, CLng(v(5)), CStr(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), CDbl(v(10)), CDbl(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), p4, CLng(v(19)), CBool(v(UBound(v))), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CLng(v(25)), CStr(v(26)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeBasicDig__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(11)
    TrimCodeBasicDig__ = VBAProject.VBT_LIB_HardIP.TrimCodeBasicDig(p1, p2, CLng(v(2)), CDbl(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), CStr(v(10)), p3, CLng(v(12)), CBool(v(UBound(v))), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PCIE_Eye_Diagram_0__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PCIE_Eye_Diagram_0__ = VBAProject.VBT_LIB_HardIP.PCIE_Eye_Diagram_0()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PCIE_Eye_Diagram_1__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PCIE_Eye_Diagram_1__ = VBAProject.VBT_LIB_HardIP.PCIE_Eye_Diagram_1()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PCIE_Eye_Diagram_2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PCIE_Eye_Diagram_2__ = VBAProject.VBT_LIB_HardIP.PCIE_Eye_Diagram_2()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PCIE_Eye_Diagram_3__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PCIE_Eye_Diagram_3__ = VBAProject.VBT_LIB_HardIP.PCIE_Eye_Diagram_3()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PCIE_Eye_Diagram_4__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PCIE_Eye_Diagram_4__ = VBAProject.VBT_LIB_HardIP.PCIE_Eye_Diagram_4()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PCIE_Eye_Diagram_5__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PCIE_Eye_Diagram_5__ = VBAProject.VBT_LIB_HardIP.PCIE_Eye_Diagram_5()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LPDPRX_Eye_Diagram__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    LPDPRX_Eye_Diagram__ = VBAProject.VBT_LIB_HardIP.LPDPRX_Eye_Diagram()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Enable_HIP_Datalog_Format__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Enable_HIP_Datalog_Format__ = VBAProject.VBT_LIB_HardIP.Enable_HIP_Datalog_Format()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Disable_HIP_Datalog_Format__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Disable_HIP_Datalog_Format__ = VBAProject.VBT_LIB_HardIP.Disable_HIP_Datalog_Format()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_SEPVM_Ref_Level_Div__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Set_SEPVM_Ref_Level_Div__ = VBAProject.VBT_LIB_HardIP.Set_SEPVM_Ref_Level_Div()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SEPVM_Ref_measurement__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    Dim p5 As New PinList
    p5.Value = v(4)
    SEPVM_Ref_measurement__ = VBAProject.VBT_LIB_HardIP.SEPVM_Ref_measurement(p1, p2, p3, p4, p5, CDbl(v(5)), CDbl(v(6)), CDbl(v(7)), CDbl(v(8)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SEPVM_Ref2_Calibration__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    SEPVM_Ref2_Calibration__ = VBAProject.VBT_LIB_HardIP.SEPVM_Ref2_Calibration(p1, p2, p3, p4, CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), CDbl(v(7)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReSet_SEPVM_Ref_Level_Div__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ReSet_SEPVM_Ref_Level_Div__ = VBAProject.VBT_LIB_HardIP.ReSet_SEPVM_Ref_Level_Div()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LDO_Calibration__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(4)
    LDO_Calibration__ = VBAProject.VBT_LIB_HardIP.LDO_Calibration(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), p2, CLng(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CDbl(v(9)), CLng(v(10)), CLng(v(11)), CStr(v(12)), CDbl(v(13)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function HIP_eFuse_Read_TMPS_Coeff__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    HIP_eFuse_Read_TMPS_Coeff__ = VBAProject.VBT_LIB_HardIP_AP.HIP_eFuse_Read_TMPS_Coeff(CStr(v(0)), CStr(v(1)), CStr(v(2)), CLng(v(3)), CBool(v(4)), CStr(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TMPS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(2)
    Dim p3 As New PinList
    p3.Value = v(7)
    Dim p4 As CalculateMethodSetup
    p4 = v(10)
    TMPS__ = VBAProject.VBT_LIB_HardIP_AP.TMPS(p1, CBool(v(1)), p2, CLng(v(3)), CLng(v(4)), CStr(v(5)), CStr(v(6)), p3, CLng(v(8)), CLng(v(9)), p4, CStr(v(11)), CStr(v(12)), CStr(v(13)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HIP_eFuse_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    HIP_eFuse_Write__ = VBAProject.VBT_LIB_HardIP_AP.HIP_eFuse_Write(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CBool(v(4)), CStr(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ADC_Trim__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    ADC_Trim__ = VBAProject.VBT_LIB_HardIP_AP.ADC_Trim(p1, CBool(v(1)), CStr(v(2)), p2, CLng(v(4)), CLng(v(5)), CStr(v(6)), CStr(v(7)), CDbl(v(8)), CStr(v(9)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MTR_Sense_Calibration_Coeff_Verification__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MTR_Sense_Calibration_Coeff_Verification__ = VBAProject.VBT_LIB_HardIP_AP.MTR_Sense_Calibration_Coeff_Verification(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Metrology_CAL_eFuse_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Metrology_CAL_eFuse_Write__ = VBAProject.VBT_LIB_HardIP_AP.Metrology_CAL_eFuse_Write(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CBool(v(8)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HIP_eFuse_Write_by_MTRGSNS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    HIP_eFuse_Write_by_MTRGSNS__ = VBAProject.VBT_LIB_HardIP_AP.HIP_eFuse_Write_by_MTRGSNS(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CBool(v(4)), CStr(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Check_MTRGSNS_25C_Fuse__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Check_MTRGSNS_25C_Fuse__ = VBAProject.VBT_LIB_HardIP_AP.Check_MTRGSNS_25C_Fuse()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MTR_REL_Fuse_Calc_Verification__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MTR_REL_Fuse_Calc_Verification__ = VBAProject.VBT_LIB_HardIP_AP.MTR_REL_Fuse_Calc_Verification(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MTR_Sense_Alignment_Calc__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MTR_Sense_Alignment_Calc__ = VBAProject.VBT_LIB_HardIP_AP.MTR_Sense_Alignment_Calc(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MTR_Sense_Calibration_Coeff_Calc__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MTR_Sense_Calibration_Coeff_Calc__ = VBAProject.VBT_LIB_HardIP_AP.MTR_Sense_Calibration_Coeff_Calc(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MTRG_t5p2a_DigSrc_Coefficient_PreCalc__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MTRG_t5p2a_DigSrc_Coefficient_PreCalc__ = VBAProject.VBT_LIB_HardIP_AP.MTRG_t5p2a_DigSrc_Coefficient_PreCalc(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MTRG_t6p3abc_DigSrc_Coefficient_PreCalc__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MTRG_t6p3abc_DigSrc_Coefficient_PreCalc__ = VBAProject.VBT_LIB_HardIP_AP.MTRG_t6p3abc_DigSrc_Coefficient_PreCalc(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DSSC_Search__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(1)
    Dim p2 As New PinList
    p2.Value = v(2)
    DSSC_Search__ = VBAProject.VBT_LIB_HardIP_AP.DSSC_Search(CStr(v(0)), p1, p2, CDbl(v(3)), CLng(v(4)), CStr(v(5)), CLng(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DSSC_Search_LDO__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(5)
    DSSC_Search_LDO__ = VBAProject.VBT_LIB_HardIP_AP.DSSC_Search_LDO(p1, p2, CStr(v(2)), CStr(v(3)), CStr(v(4)), p3, CLng(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CDbl(v(10)), CLng(v(11)), CLng(v(12)), CStr(v(13)), CDbl(v(14)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimUVI80Code_VFI_2sComplement__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(3)
    Dim p2 As New PinList
    p2.Value = v(4)
    Dim p3 As New PinList
    p3.Value = v(6)
    Dim p4 As EventSourceWithTerminationMode
    p4 = v(8)
    Dim p5 As New PinList
    p5.Value = v(12)
    Dim p6 As New PinList
    p6.Value = v(17)
    TrimUVI80Code_VFI_2sComplement__ = VBAProject.VBT_LIB_HardIP_AP.TrimUVI80Code_VFI_2sComplement(CBool(v(0)), CStr(v(1)), CStr(v(2)), p1, p2, CDbl(v(5)), p3, CStr(v(7)), p4, CDbl(v(9)), CLng(v(10)), CStr(v(11)), p5, CLng(v(13)), CLng(v(14)), CStr(v(15)), CStr(v(16)), p6, CLng(v(18)), CLng(v(19)), CStr(v(20)), CStr(v(21)), CDbl(v(22)), CLng(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CBool(v(27)), CBool(v(UBound(v))), CStr(v(29)), CBool(v(30)), CStr(v(31)), CBool(v(32)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimUVI80Code_VFI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(3)
    Dim p2 As New PinList
    p2.Value = v(5)
    Dim p3 As EventSourceWithTerminationMode
    p3 = v(7)
    Dim p4 As New PinList
    p4.Value = v(11)
    Dim p5 As New PinList
    p5.Value = v(16)
    TrimUVI80Code_VFI__ = VBAProject.VBT_LIB_HardIP_AP.TrimUVI80Code_VFI(CStr(v(0)), CStr(v(1)), CStr(v(2)), p1, CDbl(v(4)), p2, CStr(v(6)), p3, CDbl(v(8)), CLng(v(9)), CStr(v(10)), p4, CLng(v(12)), CLng(v(13)), CStr(v(14)), CStr(v(15)), p5, CLng(v(17)), CLng(v(18)), CStr(v(19)), CStr(v(20)), CDbl(v(21)), CLng(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CBool(v(26)), CStr(v(27)), CBool(v(28)), CStr(v(29)), CBool(v(30)), CStr(v(31)), CStr(v(32)), CBool(v(33)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimUVI80Code_VFI_ADC__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(2)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(5)
    Dim p4 As EventSourceWithTerminationMode
    p4 = v(7)
    Dim p5 As New PinList
    p5.Value = v(11)
    Dim p6 As New PinList
    p6.Value = v(16)
    TrimUVI80Code_VFI_ADC__ = VBAProject.VBT_LIB_HardIP_AP.TrimUVI80Code_VFI_ADC(CStr(v(0)), CStr(v(1)), p1, p2, CDbl(v(4)), p3, CStr(v(6)), p4, CDbl(v(8)), CLng(v(9)), CStr(v(10)), p5, CLng(v(12)), CLng(v(13)), CStr(v(14)), CStr(v(15)), p6, CLng(v(17)), CLng(v(18)), CStr(v(19)), CStr(v(20)), CDbl(v(21)), CLng(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CBool(v(26)), CBool(v(UBound(v))), CStr(v(28)), CBool(v(29)), CStr(v(30)), CBool(v(31)), CDbl(v(32)), CBool(v(33)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeFreq_new__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(4)
    TrimCodeFreq_new__ = VBAProject.VBT_LIB_HardIP_AP.TrimCodeFreq_new(p1, CStr(v(1)), CBool(v(2)), p2, p3, CLng(v(5)), CBool(v(6)), CBool(v(7)), CBool(v(8)), CDbl(v(9)), CDbl(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CBool(v(UBound(v))), CStr(v(18)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeFreq_new_0828__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(4)
    TrimCodeFreq_new_0828__ = VBAProject.VBT_LIB_HardIP_AP.TrimCodeFreq_new_0828(p1, CStr(v(1)), CBool(v(2)), p2, p3, CLng(v(5)), CBool(v(6)), CBool(v(7)), CBool(v(8)), CDbl(v(9)), CDbl(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CBool(v(UBound(v))), CStr(v(18)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeFreq__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(4)
    TrimCodeFreq__ = VBAProject.VBT_LIB_HardIP_AP.TrimCodeFreq(p1, CStr(v(1)), CBool(v(2)), p2, p3, CLng(v(5)), CBool(v(6)), CBool(v(7)), CBool(v(8)), CDbl(v(9)), CDbl(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CBool(v(UBound(v))), CStr(v(18)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeImpedence__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(2)
    Dim p3 As New PinList
    p3.Value = v(3)
    Dim p4 As New PinList
    p4.Value = v(5)
    TrimCodeImpedence__ = VBAProject.VBT_LIB_HardIP_AP.TrimCodeImpedence(p1, CBool(v(1)), p2, p3, CStr(v(4)), p4, CLng(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), CDbl(v(10)), CDbl(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CLng(v(17)), CLng(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), CLng(v(22)), CBool(v(23)), CBool(v(UBound(v))), CStr(v(25)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HIP_eFuse_Read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    HIP_eFuse_Read__ = VBAProject.VBT_LIB_HardIP_AP.HIP_eFuse_Read(CStr(v(0)), CStr(v(1)), CStr(v(2)), CLng(v(3)), CBool(v(4)), CStr(v(5)), CStr(v(6)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MTR_Verification_Calculate__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' MTR_Verification_Calculate__ = VBAProject.VBT_LIB_HardIP_AP.MTR_Verification_Calculate(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeFreq_New_ALG__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(3)
    Dim p3 As New PinList
    p3.Value = v(4)
    TrimCodeFreq_New_ALG__ = VBAProject.VBT_LIB_HardIP_AP.TrimCodeFreq_New_ALG(p1, CStr(v(1)), CBool(v(2)), p2, p3, CLng(v(5)), CBool(v(6)), CBool(v(7)), CBool(v(8)), CDbl(v(9)), CDbl(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CInt(v(17)), CStr(v(18)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeFreq_RunPat_and_MeasF__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(3)
    Dim p2 As New PinList
    p2.Value = v(5)
    ' TrimCodeFreq_RunPat_and_MeasF__ = VBAProject.VBT_LIB_HardIP_AP.TrimCodeFreq_RunPat_and_MeasF(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeFreq_WriteComment_DspTrimCode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' TrimCodeFreq_WriteComment_DspTrimCode__ = VBAProject.VBT_LIB_HardIP_AP.TrimCodeFreq_WriteComment_DspTrimCode(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function AMP_EYE_VT_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    AMP_EYE_VT_Setup__ = VBAProject.VBT_LIB_HardIP_Customize.AMP_EYE_VT_Setup(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CUS_AMP_SDLL_SWP_Init__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CUS_AMP_SDLL_SWP_Init__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_AMP_SDLL_SWP_Init(CLng(v(0)), CLng(v(1)), CLng(v(2)), CStr(v(3)), CStr(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CUS_VIR_MainProgram_MeasV_CalR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' CUS_VIR_MainProgram_MeasV_CalR__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_VIR_MainProgram_MeasV_CalR(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function AnalyzeCusStrToCalcR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' AnalyzeCusStrToCalcR__ = VBAProject.VBT_LIB_HardIP_Customize.AnalyzeCusStrToCalcR(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Cust_Sweep_V__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Cust_Sweep_V__ = VBAProject.VBT_LIB_HardIP_Customize.Cust_Sweep_V()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VOLH_Sweep__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    VOLH_Sweep__ = VBAProject.VBT_LIB_HardIP_Customize.VOLH_Sweep(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MTR_UVI80_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MTR_UVI80_Setup__ = VBAProject.VBT_LIB_HardIP_Customize.MTR_UVI80_Setup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CUS_DDR_Emulate_Const_Res_Loading__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As Enum_RAK
    p1 = v(4)
    ' CUS_DDR_Emulate_Const_Res_Loading__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_DDR_Emulate_Const_Res_Loading(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CUS_DDR_DCS_PrintOut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CUS_DDR_DCS_PrintOut__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_DDR_DCS_PrintOut()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MEAS_I_ABS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' MEAS_I_ABS__ = VBAProject.VBT_LIB_HardIP_Customize.MEAS_I_ABS(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CUS_RREF_Rak_Calc__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' CUS_RREF_Rak_Calc__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_RREF_Rak_Calc(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CUS_AMP_SDLL_SWP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' CUS_AMP_SDLL_SWP__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_AMP_SDLL_SWP(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ADCLK_Matrix_Loading__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ADCLK_Matrix_Loading__ = VBAProject.VBT_LIB_HardIP_Customize.ADCLK_Matrix_Loading()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Time_Measure_kit_UP1600__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(6)
    Dim p4 As New PinList
    p4.Value = v(9)
    Dim p5 As New PinList
    p5.Value = v(18)
    Time_Measure_kit_UP1600__ = VBAProject.VBT_LIB_HardIP_JitterEye.Time_Measure_kit_UP1600(p1, p2, CBool(v(2)), CBool(v(3)), CBool(v(4)), CStr(v(5)), p3, CLng(v(7)), CLng(v(8)), p4, CLng(v(10)), CLng(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), p5, CBool(v(19)), CDbl(v(20)), CDbl(v(21)), CDbl(v(22)), CDbl(v(23)), CDbl(v(24)), CDbl(v(25)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DDR_RO_Time_Measure_KIT_UP1600__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(6)
    Dim p4 As New PinList
    p4.Value = v(9)
    DDR_RO_Time_Measure_KIT_UP1600__ = VBAProject.VBT_LIB_HardIP_JitterEye.DDR_RO_Time_Measure_KIT_UP1600(p1, p2, CBool(v(2)), CBool(v(3)), CBool(v(4)), CStr(v(5)), p3, CLng(v(7)), CLng(v(8)), p4, CLng(v(10)), CLng(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CBool(v(19)), CDbl(v(20)), CDbl(v(21)), CDbl(v(22)), CDbl(v(23)), CDbl(v(24)), CDbl(v(25)), CLng(v(26)), CLng(v(27)), CStr(v(28)), CStr(v(29)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function PCIE_PI_TEST__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New Pattern
    p2.Value = v(1)
    Dim p3 As New Pattern
    p3.Value = v(2)
    Dim p4 As New Pattern
    p4.Value = v(3)
    PCIE_PI_TEST__ = VBAProject.VBT_LIB_HardIP_PCIE_PI.PCIE_PI_TEST(p1, p2, p3, p4, CInt(v(4)), CLng(v(5)), CLng(v(6)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PCIE_PI_Pat1__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PCIE_PI_Pat1__ = VBAProject.VBT_LIB_HardIP_PCIE_PI.PCIE_PI_Pat1(CStr(v(0)), CLng(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PCIE_PI_Pat2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PCIE_PI_Pat2__ = VBAProject.VBT_LIB_HardIP_PCIE_PI.PCIE_PI_Pat2(CStr(v(0)), CInt(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PCIE_PI_Pat3__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PCIE_PI_Pat3__ = VBAProject.VBT_LIB_HardIP_PCIE_PI.PCIE_PI_Pat3(CStr(v(0)), CStr(v(1)), CStr(v(2)), CInt(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PCIE_PI_Pat4__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PCIE_PI_Pat4__ = VBAProject.VBT_LIB_HardIP_PCIE_PI.PCIE_PI_Pat4(CStr(v(0)), CLng(v(1)), CLng(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PI_continuous_fai_count__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' PI_continuous_fai_count__ = VBAProject.VBT_LIB_HardIP_PCIE_PI.PI_continuous_fai_count(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EyeCount__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' EyeCount__ = VBAProject.VBT_LIB_HardIP_PCIE_PI.EyeCount(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EyeWidth__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' EyeWidth__ = VBAProject.VBT_LIB_HardIP_PCIE_PI.EyeWidth(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Protect_Mbist_Sheet__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Protect_Mbist_Sheet__ = VBAProject.VBT_LIB_MBIST.Protect_Mbist_Sheet()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UnProtect_Mbist_Sheet__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UnProtect_Mbist_Sheet__ = VBAProject.VBT_LIB_MBIST.UnProtect_Mbist_Sheet()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Mbist_SetLoopCNT_BM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_Mbist_SetLoopCNT_BM__ = VBAProject.VBT_LIB_MBIST.auto_Mbist_SetLoopCNT_BM(CStr(v(0)), CStr(v(1)), CStr(v(2)), CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Mbist_Initialize__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Mbist_Initialize__ = VBAProject.VBT_LIB_MBIST.Mbist_Initialize()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function init_MBIST_ChkList_block_loop__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    init_MBIST_ChkList_block_loop__ = VBAProject.VBT_LIB_MBIST.init_MBIST_ChkList_block_loop(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Separate_nu_char__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Separate_nu_char__ = VBAProject.VBT_LIB_MBIST.Separate_nu_char(CStr(v(0)), CLng(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function UART_read_response__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_read_response__ = VBAProject.VBT_UART_TX_Module.UART_read_response()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UART_read_response_extended__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_read_response_extended__ = VBAProject.VBT_UART_TX_Module.UART_read_response_extended()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UART_boot__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_boot__ = VBAProject.VBT_UART_TX_Module.UART_boot()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function SPI_ROM_Written_record__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_Write_SPIROM.SPI_ROM_Written_record
    SPI_ROM_Written_record__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_32M_Check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    SPIROM_32M_Check__ = VBAProject.VBT_Write_SPIROM.SPIROM_32M_Check(p1, p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    SPIROM_Continuity__ = VBAProject.VBT_Write_SPIROM.SPIROM_Continuity(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PA_WriteEnable__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' PA_WriteEnable__ = VBAProject.VBT_Write_SPIROM.PA_WriteEnable(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PA_Erase__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' PA_Erase__ = VBAProject.VBT_Write_SPIROM.PA_Erase(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BE_TimeOut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' BE_TimeOut__ = VBAProject.VBT_Write_SPIROM.BE_TimeOut(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_FL_PP_VBT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(2)
    SPIROM_FL_PP_VBT__ = VBAProject.VBT_Write_SPIROM.SPIROM_FL_PP_VBT(p1, CStr(v(1)), p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Get_RomSize_VBT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SPIROM_Get_RomSize_VBT__ = VBAProject.VBT_Write_SPIROM.SPIROM_Get_RomSize_VBT(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_FL_PP_VBT_32MByte__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(2)
    SPIROM_FL_PP_VBT_32MByte__ = VBAProject.VBT_Write_SPIROM.SPIROM_FL_PP_VBT_32MByte(p1, CStr(v(1)), p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Read_RomCode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SPIROM_Read_RomCode__ = VBAProject.VBT_Write_SPIROM.SPIROM_Read_RomCode(CStr(v(0)), CLng(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Read_RomCode_32M__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SPIROM_Read_RomCode_32M__ = VBAProject.VBT_Write_SPIROM.SPIROM_Read_RomCode_32M(CStr(v(0)), CLng(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function InitialRead_VBT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    InitialRead_VBT__ = VBAProject.VBT_Write_SPIROM.InitialRead_VBT(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function InitialRead_32M_VBT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    InitialRead_32M_VBT__ = VBAProject.VBT_Write_SPIROM.InitialRead_32M_VBT(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function InitialRead_32M_VBT_SectorErase__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    InitialRead_32M_VBT_SectorErase__ = VBAProject.VBT_Write_SPIROM.InitialRead_32M_VBT_SectorErase(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckSum_VBT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    CheckSum_VBT__ = VBAProject.VBT_Write_SPIROM.CheckSum_VBT(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckSum_32M_VBT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    CheckSum_32M_VBT__ = VBAProject.VBT_Write_SPIROM.CheckSum_32M_VBT(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckSum_32M_VBT_SectorErase__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    Dim p4 As New PinList
    p4.Value = v(3)
    CheckSum_32M_VBT_SectorErase__ = VBAProject.VBT_Write_SPIROM.CheckSum_32M_VBT_SectorErase(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Program_VBT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    SPIROM_Program_VBT__ = VBAProject.VBT_Write_SPIROM.SPIROM_Program_VBT(p1, p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Program_32M_VBT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    SPIROM_Program_32M_VBT__ = VBAProject.VBT_Write_SPIROM.SPIROM_Program_32M_VBT(p1, p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_InitRead__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(5)
    SPIROM_InitRead__ = VBAProject.VBT_Write_SPIROM.SPIROM_InitRead(p1, p2, CLng(v(2)), CStr(v(3)), CInt(v(4)), p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_CheckSum__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(4)
    SPIROM_CheckSum__ = VBAProject.VBT_Write_SPIROM.SPIROM_CheckSum(p1, p2, CStr(v(2)), CLng(v(3)), p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_InitRead_32M__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(5)
    SPIROM_InitRead_32M__ = VBAProject.VBT_Write_SPIROM.SPIROM_InitRead_32M(p1, p2, CLng(v(2)), CStr(v(3)), CInt(v(4)), p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_InitRead_32M_SectorErase__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(5)
    SPIROM_InitRead_32M_SectorErase__ = VBAProject.VBT_Write_SPIROM.SPIROM_InitRead_32M_SectorErase(p1, p2, CLng(v(2)), CStr(v(3)), CInt(v(4)), p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_CheckSum_32M__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(4)
    SPIROM_CheckSum_32M__ = VBAProject.VBT_Write_SPIROM.SPIROM_CheckSum_32M(p1, p2, CStr(v(2)), CLng(v(3)), p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_CheckSum_32M_SectorErase__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(4)
    SPIROM_CheckSum_32M_SectorErase__ = VBAProject.VBT_Write_SPIROM.SPIROM_CheckSum_32M_SectorErase(p1, p2, CStr(v(2)), CLng(v(3)), p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_InitRead_32MByte__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    SPIROM_InitRead_32MByte__ = VBAProject.VBT_Write_SPIROM.SPIROM_InitRead_32MByte(p1, p2, CLng(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_CheckSum_32MByte__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New Pattern
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    SPIROM_CheckSum_32MByte__ = VBAProject.VBT_Write_SPIROM.SPIROM_CheckSum_32MByte(p1, p2, p3, CLng(v(3)), CStr(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Read_VBT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    SPIROM_Read_VBT__ = VBAProject.VBT_Write_SPIROM.SPIROM_Read_VBT(p1, p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Read_MByte_VBT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    SPIROM_Read_MByte_VBT__ = VBAProject.VBT_Write_SPIROM.SPIROM_Read_MByte_VBT(p1, p2, CInt(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_p2p_short_Power__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.Value = v(0)
    SPIROM_p2p_short_Power__ = VBAProject.VBT_Write_SPIROM.SPIROM_p2p_short_Power(p1, CDbl(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_WaitTime__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' SPIROM_WaitTime__ = VBAProject.VBT_Write_SPIROM.SPIROM_WaitTime(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_SectorErase__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New PinList
    p2.Value = v(1)
    Dim p3 As New PinList
    p3.Value = v(2)
    SPIROM_SectorErase__ = VBAProject.VBT_Write_SPIROM.SPIROM_SectorErase(p1, p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_Footer_SPIROM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_Footer_SPIROM__ = VBAProject.VBT_Write_SPIROM.Print_Footer_SPIROM(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPI_ROM_writtten_record__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SPI_ROM_writtten_record__ = VBAProject.VBT_Write_SPIROM.SPI_ROM_writtten_record()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_Header_SPIROM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_Header_SPIROM__ = VBAProject.VBT_Write_SPIROM.Print_Header_SPIROM(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function UART_write_pmgr__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_write_pmgr__ = VBAProject.VBT_UART_RX_Module.UART_write_pmgr()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function RTOS_Command__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Command__ = VBAProject.VBT_LIB_SPI_Update.RTOS_Command(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CDbl(v(5)), CDbl(v(6)), CDbl(v(7)), CDbl(v(8)), CDbl(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_eFuse_Read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_eFuse_Read__ = VBAProject.VBT_LIB_SPI_Update.RTOS_eFuse_Read(CStr(v(0)), CStr(v(1)), CStr(v(2)), CLng(v(3)), CBool(v(4)), CStr(v(5)), CStr(v(6)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_eFuse_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_eFuse_Write__ = VBAProject.VBT_LIB_SPI_Update.RTOS_eFuse_Write(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CBool(v(4)), CStr(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_IDS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' RTOS_IDS__ = VBAProject.VBT_LIB_SPI_Update.RTOS_IDS(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Boot__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Boot__ = VBAProject.VBT_LIB_SPI_Update.RTOS_Boot(CBool(v(0)), CStr(v(1)), CBool(v(2)), CBool(v(3)), CBool(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Boot_MTRSNS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Boot_MTRSNS__ = VBAProject.VBT_LIB_SPI_Update.RTOS_Boot_MTRSNS(CBool(v(0)), CStr(v(1)), CBool(v(2)), CBool(v(3)), CBool(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Boot_CZ__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' RTOS_Boot_CZ__ = VBAProject.VBT_LIB_SPI_Update.RTOS_Boot_CZ(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_RunMetrology__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_RunMetrology__ = VBAProject.VBT_LIB_SPI_Update.RTOS_RunMetrology(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CDbl(v(9)), CDbl(v(10)), CDbl(v(11)), CDbl(v(12)), CDbl(v(13)), CDbl(v(14)), CDbl(v(15)), CStr(v(16)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Shmoo_Reboot__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' RTOS_Shmoo_Reboot__ = VBAProject.VBT_LIB_SPI_Update.RTOS_Shmoo_Reboot(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Prepoint_check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' RTOS_Prepoint_check__ = VBAProject.VBT_LIB_SPI_Update.RTOS_Prepoint_check(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_RunScenario_ORI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_RunScenario_ORI__ = VBAProject.VBT_LIB_SPI_Update.RTOS_RunScenario_ORI(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CDbl(v(6)), CDbl(v(7)), CDbl(v(8)), CDbl(v(9)), CDbl(v(10)), CStr(v(11)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_RunScenario__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_RunScenario__ = VBAProject.VBT_LIB_SPI_Update.RTOS_RunScenario(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CDbl(v(6)), CDbl(v(7)), CDbl(v(8)), CDbl(v(9)), CDbl(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CInt(v(15)), CDbl(v(16)), CBool(v(17)), CStr(v(18)), CBool(v(19)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_UART_Print__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' RTOS_UART_Print__ = VBAProject.VBT_LIB_SPI_Update.RTOS_UART_Print(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SendCmdOrg__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' SendCmdOrg__ = VBAProject.VBT_LIB_SPI_Update.SendCmdOrg(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SendCmd__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' SendCmd__ = VBAProject.VBT_LIB_SPI_Update.SendCmd(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LogDUTResponse__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' LogDUTResponse__ = VBAProject.VBT_LIB_SPI_Update.LogDUTResponse(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SendCmdOnly__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_SPI_Update.SendCmdOnly(CStr(v(0)))
    SendCmdOnly__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function WriteToOutputWindow__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' Call VBAProject.VBT_LIB_SPI_Update.WriteToOutputWindow(*One or more unsupported types in argument list or non Long/Integer return type*)
    WriteToOutputWindow__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function WriteToDatalog__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' Call VBAProject.VBT_LIB_SPI_Update.WriteToDatalog(*One or more unsupported types in argument list or non Long/Integer return type*)
    WriteToDatalog__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReloadUARTModules__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ReloadUARTModules__ = VBAProject.VBT_LIB_SPI_Update.ReloadUARTModules()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_RunScenario_mod__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' RTOS_RunScenario_mod__ = VBAProject.VBT_LIB_SPI_Update.RTOS_RunScenario_mod(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Replace_Force_cmd__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' Replace_Force_cmd__ = VBAProject.VBT_LIB_SPI_Update.Replace_Force_cmd(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Boot_and_CSW__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(1)
    RTOS_Boot_and_CSW__ = VBAProject.VBT_LIB_SPI_Update.RTOS_Boot_and_CSW(CBool(v(0)), p1, CBool(v(2)), CBool(v(3)), CBool(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SendCmd_CSW__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' SendCmd_CSW__ = VBAProject.VBT_LIB_SPI_Update.SendCmd_CSW(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_RunScenario_MTR_DOE__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_RunScenario_MTR_DOE__ = VBAProject.VBT_LIB_SPI_Update.RTOS_RunScenario_MTR_DOE(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CDbl(v(6)), CDbl(v(7)), CDbl(v(8)), CDbl(v(9)), CDbl(v(10)), CStr(v(11)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Voltage_Rampdown__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Voltage_Rampdown__ = VBAProject.VBT_LIB_SPI_Update.RTOS_Voltage_Rampdown()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Voltage_RampUp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Voltage_RampUp__ = VBAProject.VBT_LIB_SPI_Update.RTOS_Voltage_RampUp()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Decide_Switching_Bit_RTOS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ' Decide_Switching_Bit_RTOS__ = VBAProject.VBT_LIB_SPI_Update.Decide_Switching_Bit_RTOS(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Boot_Up_fail_Power_Up__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Boot_Up_fail_Power_Up__ = VBAProject.VBT_LIB_SPI_Update.RTOS_Boot_Up_fail_Power_Up()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function UARTTest_CMEM_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    UARTTest_CMEM_T__ = VBAProject.VBT_LIB_Digital_UART.UARTTest_CMEM_T(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UART_LoopBackTest__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_LoopBackTest__ = VBAProject.VBT_LIB_Digital_UART.UART_LoopBackTest()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReadCMEM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ReadCMEM__ = VBAProject.VBT_LIB_Digital_UART.ReadCMEM()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UART_read_n_byte_DSP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_read_n_byte_DSP__ = VBAProject.VBT_LIB_Digital_UART.UART_read_n_byte_DSP(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UART_write_n_byte_DSP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_write_n_byte_DSP__ = VBAProject.VBT_LIB_Digital_UART.UART_write_n_byte_DSP(CLng(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PreLoad_PA_Modules__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PreLoad_PA_Modules__ = VBAProject.VBT_LIB_Digital_UART.PreLoad_PA_Modules(CStr(v(0)), CLng(v(1)), CStr(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VaryFreq_PA_UART__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    VaryFreq_PA_UART__ = VBAProject.VBT_LIB_Digital_UART.VaryFreq_PA_UART(CStr(v(0)), CDbl(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UARTReadRegDSP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UARTReadRegDSP__ = VBAProject.VBT_LIB_Digital_UART.UARTReadRegDSP()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UARTWriteRegDSP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UARTWriteRegDSP__ = VBAProject.VBT_LIB_Digital_UART.UARTWriteRegDSP(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UARTTest_CMEM_T_Update__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    UARTTest_CMEM_T_Update__ = VBAProject.VBT_LIB_Digital_UART.UARTTest_CMEM_T_Update(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReStartFRC__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_Digital_UART.ReStartFRC
    ReStartFRC__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function initVddBinning__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    initVddBinning__ = VBAProject.VBT_LIB_VDD_Binning.initVddBinning()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PrintOut_VDD_BIN__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PrintOut_VDD_BIN__ = VBAProject.VBT_LIB_VDD_Binning.PrintOut_VDD_BIN()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_DVFM_To_GradeVDD__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_DVFM_To_GradeVDD__ = VBAProject.VBT_LIB_VDD_Binning.Read_DVFM_To_GradeVDD()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2__ = VBAProject.VBT_LIB_VDD_Binning.ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_DYNAMIC_VBIN_IDS_ZONE_to_sheet__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_DYNAMIC_VBIN_IDS_ZONE_to_sheet__ = VBAProject.VBT_LIB_VDD_Binning.Print_DYNAMIC_VBIN_IDS_ZONE_to_sheet(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UpdateDLogColumns_Bincut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UpdateDLogColumns_Bincut__ = VBAProject.VBT_LIB_VDD_Binning.UpdateDLogColumns_Bincut(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function check_IDS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    check_IDS__ = VBAProject.VBT_LIB_VDD_Binning.check_IDS()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function adjust_VddBinning__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    adjust_VddBinning__ = VBAProject.VBT_LIB_VDD_Binning.adjust_VddBinning(CBool(v(0)), CBool(v(1)), CStr(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GradeSearch_VT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As tlResultMode
    p2 = v(2)
    Dim p3 As New Pattern
    p3.Value = v(5)
    GradeSearch_VT__ = VBAProject.VBT_LIB_VDD_Binning.GradeSearch_VT(p1, CStr(v(1)), p2, CStr(v(3)), CBool(v(4)), p3, CInt(v(6)), CBool(v(UBound(v))), CStr(v(8)), CLng(v(9)), CStr(v(10)), CBool(v(11)), CStr(v(12)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Binning_Calculation__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Power_Binning_Calculation__ = VBAProject.VBT_LIB_VDD_Binning.Power_Binning_Calculation(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GradeSearch_HVCC_VT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As tlResultMode
    p2 = v(2)
    Dim p3 As New Pattern
    p3.Value = v(6)
    GradeSearch_HVCC_VT__ = VBAProject.VBT_LIB_VDD_Binning.GradeSearch_HVCC_VT(p1, CStr(v(1)), p2, CStr(v(3)), CBool(v(4)), CStr(v(5)), p3, CInt(v(7)), CBool(v(8)), CBool(v(UBound(v))), CStr(v(10)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GradeSearch_postBinCut_VT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As tlResultMode
    p2 = v(2)
    Dim p3 As New Pattern
    p3.Value = v(6)
    GradeSearch_postBinCut_VT__ = VBAProject.VBT_LIB_VDD_Binning.GradeSearch_postBinCut_VT(p1, CStr(v(1)), p2, CStr(v(3)), CBool(v(4)), CStr(v(5)), p3, CInt(v(7)), CBool(v(8)), CBool(v(UBound(v))), CBool(v(10)), CStr(v(11)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BV_Init_Datalog_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BV_Init_Datalog_Setup__ = VBAProject.VBT_LIB_VDD_Binning.BV_Init_Datalog_Setup(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Restore_BV_DataLog_SetUp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Restore_BV_DataLog_SetUp__ = VBAProject.VBT_LIB_VDD_Binning.Restore_BV_DataLog_SetUp()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_BinCut_config__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_BinCut_config__ = VBAProject.VBT_LIB_VDD_Binning.Print_BinCut_config(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_VBinResult_without_Test__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Set_VBinResult_without_Test__ = VBAProject.VBT_LIB_VDD_Binning.Set_VBinResult_without_Test(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GradeSearch_CallInstance_VT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As tlResultMode
    p1 = v(1)
    GradeSearch_CallInstance_VT__ = VBAProject.VBT_LIB_VDD_Binning.GradeSearch_CallInstance_VT(CStr(v(0)), p1, CStr(v(2)), CBool(v(3)), CStr(v(4)), CBool(v(UBound(v))), CStr(v(6)), CLng(v(7)), CStr(v(8)), CBool(v(9)), CStr(v(10)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GradeSearch_HVCC_CallInstance_VT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As tlResultMode
    p1 = v(1)
    GradeSearch_HVCC_CallInstance_VT__ = VBAProject.VBT_LIB_VDD_Binning.GradeSearch_HVCC_CallInstance_VT(CStr(v(0)), p1, CStr(v(2)), CBool(v(3)), CStr(v(4)), CBool(v(UBound(v))), CStr(v(6)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Overwrite_PassBinNum_by_ForcedBin__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Overwrite_PassBinNum_by_ForcedBin__ = VBAProject.VBT_LIB_VDD_Binning.Overwrite_PassBinNum_by_ForcedBin(CBool(v(0)), CLng(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function save_siteMask_for_MultiFSTP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    save_siteMask_for_MultiFSTP__ = VBAProject.VBT_LIB_VDD_Binning.save_siteMask_for_MultiFSTP(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function restore_siteMask_for_MultiFSTP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    restore_siteMask_for_MultiFSTP__ = VBAProject.VBT_LIB_VDD_Binning.restore_siteMask_for_MultiFSTP(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Check_flagstate_for_failflag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Check_flagstate_for_failflag__ = VBAProject.VBT_LIB_VDD_Binning.Check_flagstate_for_failflag(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function align_startStep_to_GradeVDD__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    align_startStep_to_GradeVDD__ = VBAProject.VBT_LIB_VDD_Binning.align_startStep_to_GradeVDD()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function ValidateSystemSetup_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ValidateSystemSetup_T__ = OasisXLA.VBT_ConfigCheck.ValidateSystemSetup_T(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function IGSim_Functional_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.Value = v(0)
    Dim p2 As New InterposeName
    p2.Value = v(1)
    Dim p3 As New InterposeName
    p3.Value = v(2)
    Dim p4 As New InterposeName
    p4.Value = v(3)
    Dim p5 As New InterposeName
    p5.Value = v(4)
    Dim p6 As New InterposeName
    p6.Value = v(5)
    Dim p7 As New InterposeName
    p7.Value = v(6)
    Dim p8 As PFType
    p8 = v(7)
    Dim p9 As tlResultMode
    p9 = v(8)
    Dim p10 As New PinList
    p10.Value = v(9)
    Dim p11 As New PinList
    p11.Value = v(10)
    Dim p12 As New PinList
    p12.Value = v(11)
    Dim p13 As New PinList
    p13.Value = v(12)
    Dim p14 As New PinList
    p14.Value = v(13)
    Dim p15 As New PinList
    p15.Value = v(20)
    Dim p16 As New PinList
    p16.Value = v(21)
    Dim p17 As New InterposeName
    p17.Value = v(22)
    Dim p18 As tlRelayMode
    p18 = v(24)
    Dim p19 As tlWaitVal
    p19 = v(27)
    Dim p20 As tlWaitVal
    p20 = v(28)
    Dim p21 As tlWaitVal
    p21 = v(29)
    Dim p22 As tlWaitVal
    p22 = v(30)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim p23 As tlPatConcurrentMode
    p23 = v(34)
    IGSim_Functional_T__ = OasisXLA.VBT_IGSIM_FUNCTIONAL_T.IGSim_Functional_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), p15, p16, p17, CStr(v(23)), p18, CBool(v(25)), CBool(v(26)), p19, p20, p21, p22, CBool(v(UBound(v))), CStr(v(32)), pStep, CStr(v(33)), p23, CStr(v(35)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function IGSIM_PinPMU_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New InterposeName
    p1.Value = v(1)
    Dim p2 As New InterposeName
    p2.Value = v(2)
    Dim p3 As New InterposeName
    p3.Value = v(3)
    Dim p4 As New InterposeName
    p4.Value = v(4)
    Dim p5 As New InterposeName
    p5.Value = v(5)
    Dim p6 As New InterposeName
    p6.Value = v(6)
    Dim p7 As New Pattern
    p7.Value = v(7)
    Dim p8 As New Pattern
    p8.Value = v(8)
    Dim p9 As New PinList
    p9.Value = v(10)
    Dim p10 As New PinList
    p10.Value = v(11)
    Dim p11 As New PinList
    p11.Value = v(12)
    Dim p12 As New PinList
    p12.Value = v(13)
    Dim p13 As New PinList
    p13.Value = v(14)
    Dim p14 As New PinList
    p14.Value = v(15)
    Dim p15 As tlPPMUMode
    p15 = v(16)
    Dim p16 As New FormulaArg
    p16.Value = v(18)
    Dim p17 As New FormulaArg
    p17.Value = v(19)
    Dim p18 As tlPPMURelayMode
    p18 = v(20)
    Dim p19 As New PinList
    p19.Value = v(36)
    Dim p20 As New PinList
    p20.Value = v(37)
    Dim p21 As tlWaitVal
    p21 = v(38)
    Dim p22 As tlWaitVal
    p22 = v(39)
    Dim p23 As tlWaitVal
    p23 = v(40)
    Dim p24 As tlWaitVal
    p24 = v(41)
    Dim p25 As tlPPMUMode
    p25 = v(49)
    Dim p26 As New FormulaArg
    p26.Value = v(52)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim p27 As New PinList
    p27.Value = v(53)
    Dim p28 As tlPPMUMode
    p28 = v(54)
    Dim p29 As New FormulaArg
    p29.Value = v(55)
    IGSIM_PinPMU_T__ = OasisXLA.VBT_IGSIM_PinPMU_T.IGSIM_PinPMU_T(CStr(v(0)), p1, p2, p3, p4, p5, p6, p7, p8, CStr(v(9)), p9, p10, p11, p12, p13, p14, p15, CDbl(v(17)), p16, p17, p18, CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CLng(v(32)), CBool(v(33)), CStr(v(34)), CStr(v(35)), p19, p20, p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(43)), CStr(v(44)), , CStr(v(45)), CBool(v(46)), CBool(v(47)), CBool(v(48)), p25, CStr(v(50)), CStr(v(51)), p26, pStep, p27, p28, p29, CStr(v(56)), CStr(v(57)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IGSim_PPMUMeasure__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As tlRelayMode
    p1 = v(3)
    ' IGSim_PPMUMeasure__ = OasisXLA.VBT_IGSIM_PinPMU_T.IGSim_PPMUMeasure(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































