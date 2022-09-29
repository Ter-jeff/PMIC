Attribute VB_Name = "LIB_OTP_CRC"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

'************************************************
'         AHB_GlobalVariable
'************************************************
Public g_asAhbRegNameInOtp() As String ''''20200313 add, replace g_dictCRCByAHBRegName for TTR
Public g_aslAHBData() As New SiteLong '7056
Public g_dictAHBRegToOTPDataIdx        As New Dictionary
Public g_dictAHBRegQty                 As New Dictionary
Public g_dictAHBEnumIdx                As New Dictionary
Public g_dictCRCByAHBRegName           As New Dictionary  'Set Index from gC_CRCTable
Public g_slOfflineCRC                  As New SiteLong    'Offline_CRC

Public g_arrAHBRegToOTPDataIdx() As String
'Public g_aslAHBReadData(g_iAHB_CRCSIZE)  As New SiteLong  'Run AHBRead


 
Public Function CheckHwCrc(ByRef r_slCRCDataHW As SiteLong, WaitTime As Double)
On Error GoTo ErrHandler
Dim sFuncName As String: sFuncName = "CRC_PstBrn_HWCRC"
    '___Never module burst here since it will get different values every time and burn the wrong value to OTP CRC.
    '___Cyclic Integrity Check
    Dim slCRCDataHW As New SiteLong

    If TheExec.TesterMode = testModeOffline Then
        slCRCDataHW = g_slOfflineCRC
    Else
        '___HW CRC
        '___clear OTP CRC event
        g_RegVal = &H1: AHB_WRITE "RTC_FAULT_LOG_LOG_FAULT4.FLT_OTP_CRC", g_RegVal
        '___trigger Full read CRC
        g_RegVal = &H40: AHB_WRITE "OTP_SLV_OTP_CMD", g_RegVal
        g_RegVal = &H4: AHB_WRITE "OTP_SLV_OTP_CFG1.PTM", g_RegVal
        '___Wait time MUST
        TheHdw.Wait WaitTime
        '___trigger Full read CRC HWCRC value= slCRCDataHW
        AHB_READ "OTP_SLV_OTP_CRC", slCRCDataHW
    End If
    
    r_slCRCDataHW = slCRCDataHW
    TheExec.Flow.TestLimit slCRCDataHW, 0, 255, TName:="OTP_CRC_HWCRC", Unit:=unitNone 'CRC_ReadValue_OTP__UNUSED1_0
 
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function GetSwCrc_ByOtpCat(r_tWriteOrReadREG As g_eRegWriteRead) As SiteLong
    Dim sFuncName As String: sFuncName = "GetSwCrc_ByOtpCat"
    On Error GoTo ErrHandler
    'Dim crcTestObj As New CRC8
    Dim slCRCValue As New SiteLong
    TheHdw.StartStopwatch 'Timer start

    '___Get OTP Read/Write data and compose g_aslAHBData which is to be calculated
    CreateDatabaseForSwCrc r_tWriteOrReadREG
    
    '___Calc SW CRC
    Call CalculateSwCrc(g_aslAHBData(), slCRCValue)

    Set GetSwCrc_ByOtpCat = slCRCValue
    
    Call OTP_SPT_D(" *** OTP DESIGN/SYSTEM  , Exe Time =  ") 'Timer stop
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckCrcConsistency_FW(r_bOTPReadBack As Boolean) ', OTP_doneREG_local As SiteBoolean, Select_Block As OTPBLOCK_TYPE)
    Dim sFuncName As String: sFuncName = "CheckCrcConsistency_FW"
    On Error GoTo ErrHandler
    Dim slCRCDataHW     As New SiteLong
    Dim slCRCDataSW     As New SiteLong
    Dim slCRCMatch As New SiteLong

    If r_bOTPReadBack = True Then
        CheckOtpWriteReadData
    Else
                      
        '//Check OTP SW CRC vs. HW CRC
        'MP3P:
        If TheExec.Sites.Selected.Count = 0 Then Exit Function
            
        'A).OTP SW CRC
        TheHdw.StartStopwatch 'Timer start
        TheExec.Datalog.WriteComment vbCrLf & "Calc_CRC8::==>ByOTPCat"
        slCRCDataSW = GetSwCrc_ByOtpCat(eREGWRITE)
        g_OTPData.Category(SearchOtpIdxByName(g_sOTP_CRC_BIT_REG)).Read.Value = slCRCDataSW '20190701
        Call OTP_SPT_D(" *** OTP_SW_CRC , Exe Time =  ") 'Timer stop
            
            
        'B).OTP HW CRC
        TheHdw.StartStopwatch 'Timer start
        'g_RegVal = &H1: AHB_WRITE OTP_SLV_OTP_EVENT.Addr, g_RegVal
        g_RegVal = &H1: AHB_WRITE "RTC_FAULT_LOG_LOG_FAULT4.FLT_OTP_CRC", g_RegVal
        g_RegVal = &H40: AHB_WRITE "OTP_SLV_OTP_CMD", g_RegVal
        g_RegVal = &H4: AHB_WRITE "OTP_SLV_OTP_CFG1.PTM", g_RegVal
        'If Get_SPMI_STATUS = eEnable Then TheHdw.Wait 3 * ms  'MP3P SPMI
        If TheExec.TesterMode = testModeOffline Then
            slCRCDataHW = slCRCDataSW
        Else
            AHB_READ "OTP_SLV_OTP_CRC", slCRCDataHW: g_RegVal = &H0: AHB_WRITE "OTP_SLV_OTP_CMD", g_RegVal
        End If
Call OTP_SPT_D(" *** OTP_HW_CRC , Exe Time =  ") 'Timer stop
            
            
        'C).CRC Check Result:
        'slCRCMatch = slCRCDataHW.Subtract(slCRCDataSW).Abs
        slCRCMatch = slCRCDataHW.Compare(EqualTo, slCRCDataSW)
        
        'D).Datalog
        TheExec.Flow.TestLimit slCRCDataHW, 0, 255, TName:="OTP_CRC_HWCRC", Unit:=unitNone
        TheExec.Flow.TestLimit slCRCDataSW, 0, 255, TName:="OTP_CRC_SWCRC", Unit:=unitNone
        TheExec.Flow.TestLimit slCRCMatch, -1, -1, TName:="OTP_CRC_CRCCheckResult", Unit:=unitNone
    End If
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'20190418 To get OTPData Write/Read values and put them into g_aslAHBData'
Public Function CreateDatabaseForSwCrc(Optional r_eWrtieOrReadReg As g_eRegWriteRead)
    Dim sFuncName As String: sFuncName = "CreateDatabaseForSwCrc"
    On Error GoTo ErrHandler
    
    Dim lAHBIdx, lOtpIdx  As Long
    Dim lAHBRegOtptIdx As Long
    Dim sAHBRegName As String
    'Dim lOTPBW As Long
    Dim slData As New SiteLong
    Dim slOTPData As New SiteLong
    Dim slAHBData As New SiteLong
    Dim lAHBOffset As Long
    Dim lOVWStart, lOVWEnd As Long
    
    lOVWStart = SearchOtpIdxByName("OTP__PADDING_0") + 1
    lOVWEnd = lOVWStart + g_iOvwCrcCnt - 1
    
    '___init g_aslAHBData and release memory everytime
    Erase g_aslAHBData
    ReDim g_aslAHBData(g_iAHB_CRCSIZE + g_iOvwCrcCnt)
    
    Dim sAHBRegToOtpIdxCombine As String
    Dim asOTPIdx() As String

'thehdw.StartStopwatch 'Timer start
    
    Dim writeFlag As Long
    If (r_eWrtieOrReadReg = eREGWRITE) Then
        writeFlag = 1
    ElseIf (r_eWrtieOrReadReg = eREGREAD) Then
        writeFlag = 0
    End If


    '___Cacl gSL_AHB_Data first
''''    For lAHBIdx = 0 To g_dictCRCByAHBRegName.Count - 1
''''        sAHBRegName = g_dictCRCByAHBRegName.Keys(lAHBIdx) ''''TT=1.X ms, need to change other way.

    For lAHBIdx = 0 To UBound(g_asAhbRegNameInOtp) ''''20200313 TTR
        sAHBRegName = g_asAhbRegNameInOtp(lAHBIdx)
        slAHBData = 0

        sAHBRegToOtpIdxCombine = g_dictAHBRegToOTPDataIdx(UCase(sAHBRegName))
        asOTPIdx = Split(sAHBRegToOtpIdxCombine, ",")

'        sAHBRegToOtpIdxCombine = g_arrAHBRegToOTPDataIdx(m_lAHBIdx)
'        m_asOTPIdx = Split(m_sAHBRegToOtpIdxCombine, ",")

        For lAHBRegOtptIdx = 0 To UBound(asOTPIdx)
            lOtpIdx = CLng(asOTPIdx(lAHBRegOtptIdx))
            
            lAHBOffset = g_OTPData.Category(lOtpIdx).lOtpRegOfs
            'lOTPBW = g_OTPData.Category(lOTPIdx).lBitWidth
            
            If (writeFlag) Then
                slOTPData = g_OTPData.Category(lOtpIdx).Write.Value
            Else
                slOTPData = g_OTPData.Category(lOtpIdx).Read.Value
            End If
            
            slData = slOTPData.ShiftLeft(lAHBOffset)
            slAHBData = slAHBData.Add(slData)
        Next lAHBRegOtptIdx
        g_aslAHBData(lAHBIdx) = slAHBData
    Next lAHBIdx
    
    
'Call OTP_SPT_D(" *** OTP DESIGN/SYSTEM  , Exe Time =  ") 'Timer stop
    '2019-03-08:Handle Overwrite bits for CRC
    If g_iOvwCrcCnt = 0 Then
       Exit Function
    ElseIf g_iOvwCrcCnt > 0 Then
        'ReDim g_OVW_CRC_Info(g_iOvwCrcCnt - 1)
        'For lAHBIdx = g_iAHB_CRCSIZE + 1 To g_iAHB_CRCSIZE + g_iOvwCrcCnt
        'lOtpIdx = lOVWStart + lAHBIdx + 2 - g_iAHB_CRCSIZE
         lAHBIdx = g_dictCRCByAHBRegName.Count '4442
         For lOtpIdx = lOVWStart To lOVWEnd
            slAHBData = 0
                Select Case r_eWrtieOrReadReg
                    Case eREGWRITE
                         slData = g_OTPData.Category(lOtpIdx).Write.Value
                    Case eREGREAD
                         slData = g_OTPData.Category(lOtpIdx).Read.Value
                End Select
            g_aslAHBData(lAHBIdx) = slData
            lAHBIdx = lAHBIdx + 1
        Next lOtpIdx
    End If
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CalculateSwCrc(r_aslByteList() As SiteLong, ByRef r_slCRCValue As SiteLong) As Long
    Dim sFuncName As String: sFuncName = "CalculateSwCrc"
    On Error GoTo ErrHandler
    
    Dim wByteList As New DSPWave
    Dim lListIdx As Long
    Dim alTemp() As Long
    Dim lByteListSize As Long: lByteListSize = UBound(r_aslByteList) + 1
    ReDim alTemp(lByteListSize - 1)
    
    
    For Each Site In TheExec.Sites
        wByteList.CreateConstant 0, lByteListSize, DspLong
        For lListIdx = 0 To UBound(r_aslByteList)
            alTemp(lListIdx) = r_aslByteList(lListIdx)(Site)
        Next lListIdx
        
        wByteList.Data = alTemp
    Next Site
    
    Call RunDsp.otp_CalculateCRC(wByteList, r_slCRCValue)


Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function CheckHwCrcEndurance_PTM() As Long
    Dim sFuncName As String: sFuncName = "CheckHwCrcEndurance_PTM"
    On Error GoTo ErrHandler
    Dim sParmName          As String
    Dim lModeIdx           As Long
    Dim alPTMMode(2)       As Long
    Dim aslCRCDataHW()     As New SiteLong
    Dim aslEventRTCFault() As New SiteLong
    Dim slCRCMatchAll      As New SiteLong
    Dim sFlowSheetName     As String
    Dim sDatalog           As String
    ReDim aslCRCDataHW(2)
    ReDim aslEventRTCFault(2)
    
    sFlowSheetName = Replace(TheExec.Flow.CurrentFlowSheetName, "Flow_", "")
    sDatalog = g_sOTPRevisionType & "_" & sFlowSheetName & "_"
    
    'Note:(2019-06-13)Read CRC Error
    'Avus: OTP_SLV_OTP_EVENT
    'COTA: OTP_SLV_OTP_FAULT
    '-------------------------
    If TheExec.Sites.Selected.Count = 0 Then
       TheExec.Datalog.WriteComment "*** If no Site alive, bypass " & sFuncName & " **************************************************"
       Exit Function
    End If

    '1). Do this sequence for  PTM-0,1,4 for the PTM bit
    alPTMMode(0) = 0
    alPTMMode(1) = 1
    alPTMMode(2) = 4
    
    If TheExec.TesterMode = testModeOffline Then
        For lModeIdx = 0 To UBound(alPTMMode)
            aslCRCDataHW(lModeIdx) = 128
            aslEventRTCFault(lModeIdx) = 0
        Next lModeIdx
    Else
        For lModeIdx = 0 To UBound(alPTMMode)
            'g_RegVal = &H1: AHB_WRITE OTP_SLV_OTP_EVENT.Addr, g_RegVal  'OTP_SLV_OTP_EVENT
            g_RegVal = &H1: AHB_WRITE "RTC_FAULT_LOG_LOG_FAULT4.FLT_OTP_CRC", g_RegVal
            '//'#full read CRC
            g_RegVal = &H40: AHB_WRITE "OTP_SLV_OTP_CMD", g_RegVal   '#full read CRC
            g_RegVal = alPTMMode(lModeIdx): AHB_WRITE "OTP_SLV_OTP_CFG1.PTM", g_RegVal
            
            
            'If Get_SPMI_STATUS = eEnable Then TheHdw.Wait 3 * ms
            TheHdw.Wait 10 * ms
            '//#Read CRC from chip
            AHB_READ "OTP_SLV_OTP_CRC", aslCRCDataHW(lModeIdx)
            
            '//#Read event crc , if low then CRC matches the ON chip CRC computation.
            'AHB_READ OTP_SLV_OTP_EVENT.Addr, aslEventRTCFault(lModeIdx)  'Cota:  OTP_SLV_OTP_FAULT '20190820 Need to wait Moises' feedback
            AHB_READ "RTC_FAULT_LOG_LOG_FAULT4.FLT_OTP_CRC", aslEventRTCFault(lModeIdx)
            g_RegVal = &H0: AHB_WRITE "OTP_SLV_OTP_CMD", g_RegVal
        Next lModeIdx
    End If
    

    '2). Do this sequence for  PTM-0,1,4 for the PTM bit
    slCRCMatchAll = 0
    slCRCMatchAll = aslCRCDataHW(0).Subtract(aslCRCDataHW(1)).Abs
    slCRCMatchAll = slCRCMatchAll.Add(aslCRCDataHW(0).Subtract(aslCRCDataHW(2)).Abs)
    slCRCMatchAll = slCRCMatchAll.Add(aslCRCDataHW(1).Subtract(aslCRCDataHW(2)).Abs)
    
    'mSB_CRCPTMChk = aslCRCDataHW(0).Compare(EqualTo, aslCRCDataHW(1)).LogicalAnd(aslCRCDataHW(1).Compare(EqualTo, aslCRCDataHW(2)))
    
    '//Datalog
   If g_sOTPRevisionType = "OTP_V01" Then
        TheExec.Datalog.WriteComment "<" + sFuncName + ">" & ":" & g_sOTPRevisionType & "(ECID Only)"
        For lModeIdx = 0 To UBound(alPTMMode)
            sParmName = sDatalog & "CRC_HWCRC_OTP_SLV_OTP_CRC" & "_PTM" & alPTMMode(lModeIdx)
            TheExec.Flow.TestLimit aslCRCDataHW(lModeIdx), TName:=sParmName
        Next lModeIdx
        For lModeIdx = 0 To UBound(alPTMMode)
            sParmName = sDatalog & "CRC_OTP_SLV_OTP_FAULT" & "_PTM" & alPTMMode(lModeIdx)
            TheExec.Flow.TestLimit aslEventRTCFault(lModeIdx), TName:=sParmName
        Next lModeIdx
        sParmName = sDatalog & "_" & "CRC_HWCRC_OTP_SLV_OTP_CRC" & "_PTM" & "_CHECK_ALL_ERROR"
        TheExec.Flow.TestLimit slCRCMatchAll, TName:=sParmName
   Else
        For lModeIdx = 0 To UBound(alPTMMode)
            sParmName = sDatalog & "CRC_HWCRC_OTP_SLV_OTP_CRC" & "_PTM" & alPTMMode(lModeIdx)
            TheExec.Flow.TestLimit aslCRCDataHW(lModeIdx), lowVal:=0, hiVal:=255, TName:=sParmName
        Next lModeIdx
        For lModeIdx = 0 To UBound(alPTMMode)
            sParmName = sDatalog & "CRC_OTP_SLV_OTP_FAULT" & "_PTM" & alPTMMode(lModeIdx)
            TheExec.Flow.TestLimit aslEventRTCFault(lModeIdx), lowVal:=0, hiVal:=0, TName:=sParmName
        Next lModeIdx
        sParmName = sDatalog & "_" & "CRC_HWCRC_OTP_SLV_OTP_CRC" & "_PTM" & "_CHECK_ALL_ERROR"
        TheExec.Flow.TestLimit slCRCMatchAll, lowVal:=0, hiVal:=0, TName:=sParmName
   End If
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


