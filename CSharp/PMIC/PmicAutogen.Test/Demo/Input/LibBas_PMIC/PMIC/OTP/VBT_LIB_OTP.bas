Attribute VB_Name = "VBT_LIB_OTP"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

Public g_OTPData                      As g_tOTPDataCategorySyntax
Public g_OTPRev                       As OTPRevCategorySyntax

Private m_VddLevel As String
Private m_TestName As String

Private aslOTPRevVal() As New SiteLong
Public g_lDefRealUpdate() As Long
Public gFlag_POPEnd As Boolean
Public g_DictOTPPreCheckIndex        As New Dictionary

Public Function CheckDefaultReal(Optional r_bDebugPrintLog As Boolean = True)
Dim sFuncName As String: sFuncName = "CheckDefaultReal"
On Error GoTo ErrHandler
    
    Dim lOtpIdx As Integer
    Dim sComment      As String
    Dim lDiffCnt As Long
    Dim alFromWrite() As Long
    Dim alFromMap() As Long
    Dim wWriteComareMap As New DSPWave
    If g_sOTPRevisionType = "OTP_V01" Then
        TheExec.Datalog.WriteComment "<" + sFuncName + ">" & ":SKIP OTPCheck_DefaulReal for OTP_V1(Only ECID)"
        Exit Function
    End If
    
    If TheExec.Sites.Selected.Count = 0 Then
       TheExec.Datalog.WriteComment "*** If no Site alive, bypass " & sFuncName & " **************************************************"
       Exit Function
    End If

    TheExec.Datalog.WriteComment "<" + sFuncName + ">: Check these OTP Category Item.@ InstanceName=" + TheExec.DataManager.InstanceName

    '----This site Loop only do once----
    For Each Site In TheExec.Sites
        gDW_RealDef_fromWrite.Data = g_alDefRealUpdate ''''20200313
        wWriteComareMap = gDW_RealDef_fromMAP.LogicalCompare(NotEqualTo, gDW_RealDef_fromWrite)
        lDiffCnt = wWriteComareMap.CalcSum
        
        If r_bDebugPrintLog Then
            If lDiffCnt = 0 Then
                ''TheExec.Datalog.WriteComment "Site" & CStr(Site) & " registers are all setwrite."
                TheExec.Datalog.WriteComment "All registers are all setwrite."
            Else
                
                alFromWrite = gDW_RealDef_fromWrite.Data
                alFromMap = gDW_RealDef_fromMAP.Data
                
                For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
                    With g_OTPData.Category(lOtpIdx)
                        If alFromMap(lOtpIdx) > alFromWrite(lOtpIdx) Then '1(real) > 0(default)
                            sComment = "(Test item should be enabled if OTP had been set to REAL.)"
                            TheExec.Datalog.WriteComment "WARNING:" + FormatLog(.sOtpRegisterName, -(g_lOTPCateNameMaxLen + 2)) + " : " + FormatLog("NeedToUpdateRealValue", -25) & FormatLog(sComment, -25)
                        ElseIf alFromMap(lOtpIdx) < alFromWrite(lOtpIdx) Then '0(default) < 1(real)
                            sComment = "Default value has been setWrite, Please make sure this register doesn't need to be updated to 'real' "
                            TheExec.Datalog.WriteComment "WARNING:" + FormatLog(.sOtpRegisterName, -(g_lOTPCateNameMaxLen + 2)) + " : " + FormatLog("<Defalt Value>", -25) & FormatLog(sComment, -25)
                        End If
                    End With
                Next lOtpIdx
            End If
           TheExec.Datalog.WriteComment "Total WARNING Items = " & CStr(lDiffCnt) & vbCrLf
        End If

        Exit For
    Next Site
    
    '___DATALOG:
    Dim sTName As String
    Dim iHiLimit As Integer
    sTName = "OTPCheck-NeedToUpdateRealValue-WARNING-ITEMS"
    If TheExec.EnableWord("OTP_V1_LPC") = True Or TheExec.EnableWord("OTP_V1_LPD") = True Then
        iHiLimit = 1   'only trim OTP_ACORE_TRIM_VMAIN_POR_WARN_TRIM_690 / OTP_RTC_LDO_BG_TRIM_4438  in OTP_V1_LPC and OTP_V1_LPD
    Else
        iHiLimit = 3
        TheExec.Datalog.WriteComment "only trim OTP_ACORE_TRIM_VMAIN_POR_WARN_TRIM_690 / OTP_RTC_LDO_BG_TRIM_4438  in OTP_V1_LPC and OTP_V1_LPD"
    End If
    TheExec.Flow.TestLimit lDiffCnt, 0, iHiLimit, TName:=sTName '1 for CRC(real)

Exit_Function:
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function InitializeOtp()
Dim sFuncName As String: sFuncName = "InitializeOtp"
On Error GoTo ErrHandler

    g_sCurrentJobName = TheExec.CurrentJob
    g_bTTR_ALL = TheExec.EnableWord("A_Enable_TTR ") ''''20200313
    
''*******For OTP *************************
    '------------------------------------------------------------------------
    '___OTP/AHB Initialize
    If g_DictOTPNameIndex.Count = 0 Then
        '___1.1 parse_OTP Register Map and partial AHB Register Map
        '___1.2 Create dictionary according to OTP_register_name (g_DictOTPNameIndex),
            'Which is used for GetOTPDataIndex, to look up the index by OTP_register_name
        '___1.3 Get the OTP version from the selected enable word with OTP version info
        Call InitializeOtpTable

        '___2.1 Create g_dictCRCByAHBRegName Dictionary, used in SW CRC
        Call CreateDictionary_ForCalcSwCrc

        '___2.2 Create g_dictAHBRegToOTPDataIdx Dictionary, used in SW CRC
        Call CreateDictionary_Ahb2OtpIdxs

        '___3. Set the debug flags, it repeats in InitializeOtp.
        '___So for TTR purpose, could keep only one here.
        'Call InitializeOtpGlobalFlag '''20200313, it's put outside if...endif

        '___20200313 for AHB New Method by OTPData Structure
        '=========================================================
        '2020/03/09
        Call add_AHBReg_into_OTPData ''''was init_parse_AHB_Table
        Call init_AHBEnumIdx_to_OTPIndexDict
        '=========================================================
        
        '___Define OTP Chip names here
        Call DefineChipID
        '___Define OTP REV names here
        Call DefineOtpRevName
        
    End If
    '------------------------------------------------------------------------
    
     '___(OTP_CRC init look up table) '20200402 Need to do it every time
    '___Polynomial is 1_1100_1111 = 0x1CF = x^8 + x^7 + x^6 + x^3 + x^2 + x^1 + 1
    Call RunDsp.otp_Initialize_crc8(&HCF, &H0)
    
    Call InitializeOtpDataElement
    
    ''''20200313[MUST] after InitializeOtpDataElement()
    '___Update Write Data & Global write DspWave based on OTP version
    Call InitializeOtpVersion

    ''*******End of for OTP *************************

    '___OTP_Enable and OTP-FW control
    '___Put here again to prevent user revise the flags w.o. the validation or forcing stopping TP
    Call InitializeOtpGlobalFlag
    
    
    g_lDebugDumpCnt = 0 'Reset the counting of dump functions
    
    g_sLotID = "000000"
    g_lWaferID = 0
    
    '20190327 OTP-DSP
    gD_slOTP_ADDR_BW = g_iOTP_ADDR_BW
    gD_slOTP_REGDATA_BW = g_iOTP_REGDATA_BW
    'thehdw.DSP.ExecutionMode = tlDSPModeHostDebug

    '___20200313 add comments, especially IEDA data
    ''''[Notice] init all Existing sites to avoid unexpected cases
    For Each Site In TheExec.Sites.Existing
        g_slXCoord = -32768
        g_slYCoord = -32768

        g_svOTP_LotID = ""
        g_slOTP_WaferId = 0
        g_slOTP_XCoord = -32768
        g_slOTP_YCoord = -32768
        g_slOTP_Rev = -1
        
        '___Reset site boolean
        g_sbOtped = False
        g_sbOtpedECID = False
        g_sbOtpedPGM = False
        g_sbOtpedCRC = False
        
    Next Site
    '___Iniitail  OTP Version
    Erase aslOTPRevVal
    '___Init DefReal Array
    Erase g_alDefRealUpdate
    
    '___Create g_DictOTPPreCheckIndex, used in Auto_AHB_OTP_Write/Read CheckByCondition
    '___Init to check default/real and otp owner, put here so user could enable debug enable anytime and create the dictionary
    If TheExec.EnableWord("OTP_cmpAHB") = True Then
        Call CreateDictionary_ForCompAhbOtp
    End If
    
    '___Datalog the OTP_register_map parsing result/OTPAddToDataIndexCreate/OTPRev/OTPRevDataUpdate/AHB_Init
    Call ShowOtpInitialLog 'rename from auto_OTP_Result_Initialized to auto_OTP_Init_Result
    'theexec.Flow.TestLimit 1, 1, 1, TName:="auto_OTP_Result_Initialized" 'no neccessary 20190318
    Call SetOtpOvwSize
    ReDim g_aslAHBData(g_iAHB_CRCSIZE + g_iOvwCrcCnt)
    ReDim g_alDefRealUpdate(g_Total_OTP - 1)
    

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ReadWaferDataToCat()
'___Get ECID info and examine the ECID DataMapping; e.g.N98G17,02 with (X,Y)=(5,6)
'___1. SetWriteDecimal for ECID (1-1 LotID, 1-2 WaferID, 1-3 X,Y Coord.)
'___2. SetWriteDecimal for TPVersion and SVN Version
    Dim sFuncName As String: sFuncName = "ReadWaferDataToCat"
    On Error GoTo ErrHandler
    Dim sLotID As String
    Dim iLocDash As Integer
    Dim lCh2to6Idx As Long, lIdx As Long, lIdxj As Long, lIdxk As Long
    Dim sCh1st As String
    Dim sCh2to6 As String
    Dim iAscVal As Integer
    Dim sComment As String
    Dim iChkVal As Integer
    Dim sWaferID As String
    Dim sDeidBinStrM As String
    Dim lLen As Long
    Dim lDecimal As Long
    Dim lBinVal As Long
    Dim sBinStr As String
    Dim lChipIdIdx As Long
    
    '___[1-1 LotID]
    '___Examine 1st charater in LotID (e.g. "N"98G17)
    '___Latin small letter a is 97, b is 98, and so on until z is 122.
    
    
    '___Offline simulation
    If (TheExec.TesterMode = testModeOffline) Then
        If sLotID = "" Then sLotID = "N99G99" '"N98G17"
        sLotID = sLotID + "-" + Mid(Format(sWaferID, "00"), 1, 2)
    Else
        sLotID = TheExec.Datalog.Setup.LotSetup.LotID
        sWaferID = Trim(CStr(TheExec.Datalog.Setup.WaferSetup.ID))
    End If
    If (sLotID = "") Then sLotID = "000000"

    iLocDash = InStr(1, sLotID, "-")

    If iLocDash <> 0 Then
        g_sLotID = Mid(sLotID, 1, iLocDash - 1)
    Else
        g_sLotID = sLotID
    End If
    
    If g_bEnWrdOTPFTProg = False Then
    '___Exame first character
    sCh1st = Mid(g_sLotID, 1, 1)
    iAscVal = Asc(LCase(sCh1st))
    If (iAscVal < 97 Or iAscVal > 122) Then
        iChkVal = 0 'Fail
        sComment = "First Character of Prober LotID (" + UCase(sCh1st) + ") is not [A-Z]."
        TheExec.Datalog.WriteComment sComment
    Else
        iChkVal = 1 'Pass
        If (Len(g_sLotID) <> 6) Then 'check length
            sComment = "Character Numbers of Prober LotID (" + UCase(g_sLotID) + ") is NOT Six Characters."
            TheExec.Datalog.WriteComment sComment
            iChkVal = 0 'Fail
        Else
            '___Examine the charactor 2~6 in LotID (e.g. N"98G17")
            'Digit 0~9 is from 48 to 57

            For lCh2to6Idx = 2 To 6  ''''EcidCharPerLotId=6
                sCh2to6 = Mid(g_sLotID, lCh2to6Idx, 1)
                iAscVal = Asc(LCase(sCh2to6))

                If iAscVal < 97 Or iAscVal > 122 Then    'a=97 and z=122 in ANSI character '(OTP_template) redundant code
                    If iAscVal < 48 Or iAscVal > 57 Then ''0'=48 and '9'=57 in ANSI character
                        iChkVal = 0  'Fail
                        sComment = "Second-to-Sixth Characters of Prober LotID (" + UCase(g_sLotID) + ") are not [A-Z] or [0-9]."
                        TheExec.Datalog.WriteComment sComment
                        Exit For
                    Else
                        iChkVal = 1 'Pass
                    End If
                Else
                End If
            Next lCh2to6Idx
        End If
    End If
    

        TheExec.Flow.TestLimit iChkVal, 1, 1, TName:="Prober_LotID", PinName:=g_sLotID, formatStr:="%.0f"
        TheExec.Datalog.WriteComment "If Fail, Please key in LotID, or enable 'OTP_FTPorg' "
    Else
        TheExec.Datalog.WriteComment "FT Starge dose not have prober to read LotID"
    End If

    If g_bEnWrdOTPFTProg = False Then  'CP only, FT will skip this whole process(waferID and x,y)
        '___[1-2 WaferID]
        '___Examine wafer ID
        If (TheExec.TesterMode = testModeOffline) Then '___Offline simulation
            If sWaferID <> "" Then
                g_lWaferID = TheExec.Datalog.Setup.WaferSetup.ID
                If (IsNumeric(CStr(g_lWaferID)) = False) Then
                    sComment = "Prober WaferID (" + CStr(g_lWaferID) + ") is NOT numeric, set it to 25 (psudo wafer id)."
                    TheExec.Datalog.WriteComment sComment
                    g_lWaferID = 25
                End If
            Else
                g_lWaferID = 2
                TheExec.Datalog.WriteComment vbTab & "<Offline simulation> Set WaferID to 02 (pseudo wafer id)"
            End If
        Else '___Online
                If sWaferID <> "" Then
                g_lWaferID = TheExec.Datalog.Setup.WaferSetup.ID
                    If (IsNumeric(CStr(g_lWaferID)) = False) Then
                        sComment = "Prober WaferID (" + CStr(g_lWaferID) + ") is NOT numeric, set it to 0."
                        TheExec.Datalog.WriteComment sComment
                        g_lWaferID = 0
                    End If
                Else
                    g_lWaferID = 0
                End If
        End If
        '___Examine the waferID is in range 1~25
        'Syntax Check WaferID of the prober, offline is assumed as 02.
        If (g_lWaferID < 1 Or g_lWaferID > 25) Then 'range 1...25
            sComment = "Prober WaferID (" + CStr(g_lWaferID) + ") is out of the range [1...25]."
            TheExec.Datalog.WriteComment sComment
        End If
        
        If TheExec.EnableWord("OTP_Production") = True Then
            TheExec.Flow.TestLimit g_lWaferID, 1, 25, TName:="Prober_WaferID", formatStr:="%.0f"
        Else
            TheExec.Flow.TestLimit g_lWaferID, 1, 25, lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, TName:="Prober_WaferID", formatStr:="%.0f"
        End If
        
        '___[1-3 X,Y]
        '___Examine X,Y Coordination
        For Each Site In TheExec.Sites
            If (TheExec.TesterMode = testModeOffline) Then '___Offline simulation, set x,y as 15,6
                If (g_slXCoord(Site) = -32768 Or g_slYCoord(Site) = -32768) Then
                    Call SetXY(15, 6) 'set a pseudo XY coordinate
                    TheExec.Datalog.WriteComment vbTab & "<WARNING> Call setXY(15, 6) (pseudo XY_Coordinate)"
                End If
            End If
            
            '___Online
            g_slXCoord(Site) = TheExec.Datalog.Setup.WaferSetup.GetXCoord(Site)
            g_slYCoord(Site) = TheExec.Datalog.Setup.WaferSetup.GetYCoord(Site)
            
            '___Examine the x,y is within the limits
            ''''Syntax Check XY of the prober
            If (g_slXCoord(Site) < g_iXCOORD_LOW_LMT Or g_slXCoord(Site) > g_iXCOORD_HI_LMT) Then
                sComment = "Prober X_Coordinate (" + CStr(g_slXCoord(Site)) + ") is out of the range [" + CStr(g_iXCOORD_LOW_LMT) + "..." + CStr(g_iXCOORD_HI_LMT) + "]."
                TheExec.Datalog.WriteComment sComment
            End If
            If (g_slYCoord(Site) < g_iYCOORD_LOW_LMT Or g_slYCoord(Site) > g_iYCOORD_HI_LMT) Then
                sComment = "Prober Y_Coordinate (" + CStr(g_slYCoord(Site)) + ") is out of the range [" + CStr(g_iYCOORD_LOW_LMT) + "..." + CStr(g_iYCOORD_HI_LMT) + "]."
                TheExec.Datalog.WriteComment sComment
            End If
        Next Site
        
        TheExec.Flow.TestLimit g_slXCoord, g_iXCOORD_LOW_LMT, g_iXCOORD_HI_LMT, TName:="Prober_X", formatStr:="%.0f"
        TheExec.Flow.TestLimit g_slYCoord, g_iYCOORD_LOW_LMT, g_iYCOORD_HI_LMT, TName:="Prober_Y", formatStr:="%.0f"
        'theexec.Flow.TestLimit g_slXCoord, g_iXCOORD_LOW_LMT, g_iXCOORD_HI_LMT, lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, TName:="Prober_X", formatStr:="%.0f"
        'theexec.Flow.TestLimit g_slYCoord, g_iYCOORD_LOW_LMT, g_iYCOORD_HI_LMT, lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, TName:="Prober_Y", formatStr:="%.0f"

    End If 'g_bEnWrdOTPFTProg = False


    '___Print Out the summarized Prober's Information
    TheExec.Datalog.WriteComment vbCrLf & sFuncName + ":"
    TheExec.Datalog.WriteComment "  Lot ID = " + g_sLotID
    TheExec.Datalog.WriteComment "Wafer ID = " + CStr(g_lWaferID)
    TheExec.Datalog.WriteComment "----------------------------------------------------"
    For Each Site In TheExec.Sites
        'TheExec.Datalog.WriteComment "Lot ID (Site " + CStr(Site) + ")= " + g_sLotID & "/" & Chr(9) & "Wafer ID (Site " + CStr(Site) + ")= " + CStr(g_lWaferID)
        If g_bEnWrdOTPFTProg = False Then
            TheExec.Datalog.WriteComment "X coor (Site " + CStr(Site) + ")= " + CStr(g_slXCoord(Site)) & Chr(9) & "/" & Chr(9) & "Y coor (Site " + CStr(Site) + ")= " + CStr(g_slYCoord(Site))
            'TheExec.Datalog.WriteComment "X coor (Site " + CStr(Site) + ")= " + CStr(g_slXCoord(Site))
            'TheExec.Datalog.WriteComment "Y coor (Site " + CStr(Site) + ")= " + CStr(g_slYCoord(Site))
        End If
        TheExec.Datalog.WriteComment "----------------------------------------------------"
        If g_bEnWrdOTPFTProg = True Then Exit For
        'gD_wDEIDPGMBits.Clear
    Next Site

    '''''20170116 Composite OTP_Reg(0)...(7) Data Array for Device ID (DEID)

    
        ''''---------------------------------------
        '''For Shmoo Datalog
        HramLotId = g_sLotID
        HramWaferId = g_lWaferID
'        XCoord = g_slXCoord
'        YCoord = g_slYCoord
        ''''---------------------------------------
    
        '___Write to PRR-Part_TEXT in STDF
        '___Convert ECID from decimal to binary(64 bits) and hex (16) (OTP_template)
    For Each Site In TheExec.Sites
        Dim sDeidBinStrL As String
        Dim sDeidHexStrL As String
        Dim sDEIDActStrM As String
    
        sDeidBinStrM = ConvertECID2Bin(g_sLotID, g_lWaferID, g_slXCoord(Site), g_slYCoord(Site), 0)
    
    
        'Reserved code, not in used
        If (1 = 0) Then
            sDeidBinStrL = StrReverse(sDeidBinStrM)
            sDeidHexStrL = ConvertFormat_Bin2Hex(sDeidBinStrL, 16)
        End If
    
        sDEIDActStrM = g_sLotID + Format(g_lWaferID, "00") + Format(g_slXCoord(Site), "00") + Format(g_slYCoord(Site), "00") '+ "_" + g_sDftType(Site) '20170817

        If g_bEnWrdOTPFTProg = False Then 'CP only
            If Not (g_bTTR_ALL) Then TheExec.Datalog.WriteComment "Site(" & Site & ") Write DEID ActStrM = " + sDEIDActStrM + " [ LotID + WaferID + XCoord + YCoord + DFTType ] "
            Call TheExec.Datalog.Setup.SetPRRPartInfo(tl_SelectSite, Site, , sDEIDActStrM) ' + "_" & g_sDftType)
        End If
    
        lLen = Len(sDeidBinStrM)
    
        gD_wDEIDPGMBits.CreateConstant 0, lLen, DspLong

        For lIdx = 1 To lLen
            gD_wDEIDPGMBits(Site).Element(lIdx - 1) = CLng(Mid(sDeidBinStrM, lIdx, 1))
        Next lIdx
        
        '___EDID Encode
        'per 8 bits to convert to Long value (lDecimal)
        Dim alTemp() As Long
        Dim wChipIdReg As New DSPWave
        Call otp_allotECID(wChipIdReg)
        If True Then
            If (True) Then '___20200313, faster, need to verify 'Janet need check (ok)
                Dim tmpArr() As Long
                tmpArr = wChipIdReg.Data
                For lIdx = 0 To 7 Step 1
                    lDecimal = tmpArr(lIdx)
                    g_aslOTPChipReg(lIdx)(Site) = lDecimal
                    sBinStr = ConvertFormat_Dec2Bin(lDecimal, 8, alTemp)
                    TheExec.Datalog.WriteComment "Site(" & Site & ") g_aslOTPChipReg(" & lIdx & ") = " & FormatLog(lDecimal, 4) & " [" + sBinStr + "]"
                Next lIdx
            Else
                For lIdx = 0 To 7 Step 1
                    lDecimal = wChipIdReg(Site).ElementLite(lIdx)
                    g_aslOTPChipReg(lIdx)(Site) = lDecimal
                    sBinStr = ConvertFormat_Dec2Bin(lDecimal, 8, alTemp)
                    TheExec.Datalog.WriteComment "Site(" & Site & ") g_aslOTPChipReg(" & lIdx & ") = " & FormatLog(lDecimal, 4) & " [" + sBinStr + "]"
                Next lIdx
            End If

        Else
            '20200313, should be verified if using Wave.Data as Array then faster.
            For lIdx = 0 To 7 Step 1
               lDecimal = 0
               lBinVal = 0
               sBinStr = ""
               lDecimal = wChipIdReg(Site).ElementLite(lIdx)
    
               For lIdxj = 0 To 7
                   lIdxk = lIdxj + lIdx * 8
                   lBinVal = gD_wDEIDPGMBits(Site).ElementLite(lIdxk)
                   'lDecimal = lDecimal + lBinVal * (2 ^ lIdxj)
                   
                   sBinStr = CStr(lBinVal) + sBinStr
               Next lIdxj
               g_aslOTPChipReg(lIdx)(Site) = lDecimal
               TheExec.Datalog.WriteComment "Site(" & Site & ") g_aslOTPChipReg(" & lIdx & ") = " & FormatLog(lDecimal, 4) & " [" + sBinStr + "]"
            Next lIdx
        End If
    Next Site
    
    For lChipIdIdx = 0 To UBound(g_asChipIDName)
        Call auto_OTPCategory_SetWriteDecimal(g_asChipIDName(lChipIdIdx), g_aslOTPChipReg(lChipIdIdx))
    Next lChipIdIdx
     
    '___SetWrite the OTP_Rev
    Call CheckNSetWriteTPVersion
    '___Decode the OTP_Rev and datalog out
    Call CheckNGetWriteOTPVersion

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''Special case in MP5T
'Public Function Auto_Decode_OTPREG_DEID_PreDecode(Reg_read() As SiteLong)
'Dim sFuncName As String: sFuncName = "Auto_Decode_OTPREG_DEID_PreDecode"
'On Error GoTo ErrHandler
'
'    Dim lChipIdIdx As Long
'    Dim lIdx As Integer
'    Dim sDEIDActStrM As String
'    Dim sDeidBinStrM As String
'    Dim svDeidBinStrM As New SiteVariant  'Claire add  20170926
'    Dim lDecimal As Long
'    Dim sBinStr As String
'    Dim alBinArr() As Long
'    Dim sLodIDBinStr As String
'    Dim sWfIDBinStr As String
'    Dim sXcoordBinStr As String
'    Dim sYcoordBinStr As String
'    Dim sOtpRevBinStr As String
'
'
'    TheExec.Datalog.WriteComment ""
'    TheExec.Datalog.WriteComment "----------------------------------------------------"
'    TheExec.Datalog.WriteComment sFuncName + ":"
'
'    For Each Site In TheExec.Sites
'
'        ''''per 8 OTP_Reg() to convert to binary string
'        lDecimal = 0
'        sDeidBinStrM = ""
'        TheExec.Datalog.WriteComment "----------------------------------------------------"
'        For lChipIdIdx = 0 To UBound(g_asChipIDName)
'            sBinStr = ""
'            lDecimal = Reg_read(lChipIdIdx)(Site) ''''should be Reg_Read
'            sBinStr = ConvertFormat_Dec2Bin(lDecimal, 8, alBinArr) ''''per Register 8 bits
'            sDeidBinStrM = sDeidBinStrM + StrReverse(sBinStr) 'Claire
'            TheExec.Datalog.WriteComment "Site(" & Site & ") Reg_Read(" & lChipIdIdx & ") = " & FormatLog(lDecimal, 4) & " [" + sBinStr + "]" & " , &H" & Right("00" & Hex(lDecimal), 2)
'        Next lChipIdIdx
'
'        'KC ECID Full 64bit string for control+F to work
'        TheExec.Datalog.WriteComment "Site(" & Site & ") ECID := " & "&H " & Right("00" & Hex(Reg_read(0)(Site)), 2) & Right("00" & Hex(Reg_read(1)(Site)), 2) & Right("00" & Hex(Reg_read(2)(Site)), 2) & Right("00" & Hex(Reg_read(3)(Site)), 2) & Right("00" & Hex(Reg_read(4)(Site)), 2) & Right("00" & Hex(Reg_read(5)(Site)), 2) & Right("00" & Hex(Reg_read(6)(Site)), 2) & Right("00" & Hex(Reg_read(7)(Site)), 2)
'
'        TheExec.Datalog.WriteComment "Site(" & Site & ") Read DEID BinStrM = " + sDeidBinStrM + " [bit0...bit63][MSB...LSB]"
'        TheExec.Datalog.WriteComment "----------------------------------------------------"
'
'        sLodIDBinStr = Mid(sDeidBinStrM, g_iLOTID_BITS_START + 1, g_iLOTID_BITS_BW)
'        sWfIDBinStr = Mid(sDeidBinStrM, g_iWFID_BITS_START + 1, g_iWFID_BITS_BW)
'        sXcoordBinStr = Mid(sDeidBinStrM, g_iXCOORD_BITS_START + 1, g_iXCOORD_BITS_BW)
'        sYcoordBinStr = Mid(sDeidBinStrM, g_iYCOORD_BITS_START + 1, g_iYCOORD_BITS_BW)
'        sOtpRevBinStr = Mid(sDeidBinStrM, g_iOTPREV_BITS_START + 1, g_iOTPREV_BITS_BW)
'
'        g_svOTP_LotID = ""
'        For lIdx = 1 To 6
'            g_svOTP_LotID = g_svOTP_LotID + ConvertLotIdBin2Letter(Mid(sLodIDBinStr, 1 + (lIdx - 1) * 6, 6))
'        Next lIdx
'
'        g_slOTP_WaferId = CLng(ConvertFormat_Bin2Dec(sWfIDBinStr))
'        g_slOTP_XCoord = CLng(ConvertFormat_Bin2Dec(sXcoordBinStr))
'        g_slOTP_YCoord = CLng(ConvertFormat_Bin2Dec(sYcoordBinStr))
'        'g_slOTP_Rev = CLng(auto_OTP_binStr2Dec(sOtpRevBinStr))
'        svDeidBinStrM = sDeidBinStrM
'        TheExec.Datalog.WriteComment "Site(" & Site & ") OTP Read LotID   = " + g_svOTP_LotID
'        TheExec.Datalog.WriteComment "Site(" & Site & ") OTP Read WaferID = " + CStr(g_slOTP_WaferId)
'        TheExec.Datalog.WriteComment "Site(" & Site & ") OTP Read X_Coord = " + CStr(g_slOTP_XCoord)
'        TheExec.Datalog.WriteComment "Site(" & Site & ") OTP Read Y_Coord = " + CStr(g_slOTP_YCoord)
'        'TheExec.Datalog.WriteComment "Site(" & Site & ") OTP Read OTP_Rev = " + CStr(g_slOTP_Rev)
'        TheExec.Datalog.WriteComment "----------------------------------------------------"
'        TheExec.Datalog.WriteComment ""
'
'    Next Site
'
'
'
'    Dim slDieType As New SiteLong
'    Dim sDftType As String
'    Dim sTestName As String
'    Dim lTestNumber As Long
'    sDftType = "POR"
'
'
'    For Each Site In TheExec.Sites
'        g_slXCoord(Site) = g_slOTP_XCoord(Site)
'        g_slYCoord(Site) = g_slOTP_YCoord(Site)
'        If g_slXCoord(Site) Mod 6 = 5 Then
'            If g_slYCoord(Site) Mod 4 = 2 Then
'                TheExec.Sites.Item(Site).FlagState("F_C17_DFT1") = logicTrue
'                TheExec.Datalog.WriteComment "DIE TYPE is DFT1"
'                sDftType = "DFT1"
'                slDieType(Site) = 10  'Bin10
'            ElseIf g_slYCoord(Site) Mod 4 = 3 Then
'                TheExec.Sites.Item(Site).FlagState("F_C18_DFT2") = logicTrue
'                TheExec.Datalog.WriteComment "DIE TYPE is DFT2"
'                sDftType = "DFT2"
'                slDieType(Site) = 11  'Bin11
'            ElseIf g_slYCoord(Site) Mod 4 = 0 Then
'                TheExec.Sites.Item(Site).FlagState("F_C19_DFT3") = logicTrue
'                TheExec.Datalog.WriteComment "DIE TYPE is DFT3"
'                sDftType = "DFT3"
'                slDieType(Site) = 12  'Bin12
'            ElseIf g_slYCoord(Site) Mod 4 = 1 Then
'                TheExec.Sites.Item(Site).FlagState("F_C20_DFT4") = logicTrue
'                TheExec.Datalog.WriteComment "DIE TYPE is DFT4"
'                sDftType = "DFT4"
'                slDieType(Site) = 13  'Bin13
'            End If
'        ElseIf g_slXCoord(Site) Mod 6 = 0 Then
'            If g_slYCoord(Site) Mod 4 = 2 Then
'                TheExec.Sites.Item(Site).FlagState("F_C21_DFT5") = logicTrue
'                TheExec.Datalog.WriteComment "DIE TYPE is DFT5"
'                sDftType = "DFT5"
'                slDieType = 14  'Bin14
'            ElseIf g_slYCoord(Site) Mod 4 = 0 Then
'                TheExec.Sites.Item(Site).FlagState("F_C22_DFT6") = logicTrue
'                TheExec.Datalog.WriteComment "DIE TYPE is DFT6"
'                sDftType = "DFT6"
'                slDieType = 15  'Bin15
'            End If
'        Else
'            TheExec.Datalog.WriteComment "DIE TYPE is POR"
'            slDieType(Site) = 1  'Bin1
'            sDftType = "POR"
'        End If ' end of Device X check
'
'
'    '    g_sDftType = Split(sDftType, ",")
'        g_sDftType = sDftType
'
'
'        lTestNumber = TheExec.Sites(0).TestNumber
'
'        If g_sDftType = "POR" Then
'        ElseIf g_sDftType = "DFT1" Then
'            lTestNumber = lTestNumber + 1
'        ElseIf g_sDftType = "DFT2" Then
'            lTestNumber = lTestNumber + 2
'        ElseIf g_sDftType = "DFT3" Then
'            lTestNumber = lTestNumber + 3
'        ElseIf g_sDftType = "DFT4" Then
'            lTestNumber = lTestNumber + 4
'        ElseIf g_sDftType = "DFT5" Then
'            lTestNumber = lTestNumber + 5
'        ElseIf g_sDftType = "DFT6" Then
'            lTestNumber = lTestNumber + 6
'        End If
'
'
'        TheExec.Sites(Site).TestNumber = lTestNumber
'
'
'        sTestName = TNameCombine("SPA", "Device", "Type", g_sDftType, , , TName_NonTrimItem, , , TName_ToggleDTB)
'        Call TheExec.Flow.TestLimit(ResultVal:=0, ForceResults:=tlForceNone, Unit:=unitNone, TName:=sTestName, lowVal:=0, hiVal:=0, formatStr:="%.0f")
'
'        sDEIDActStrM = g_sLotID + Format(g_lWaferID, "00") + Format(g_slXCoord(Site), "00") + Format(g_slYCoord(Site), "00") '+ "_" + g_sDftType(Site) '20170817
'        Call TheExec.Datalog.Setup.SetPRRPartInfo(tl_SelectSite, Site, , sDEIDActStrM + "_" & g_sDftType)
'
'    Next Site
'
'
'
'    Call TheExec.Flow.TestLimit(ResultVal:=sDftType, ForceResults:=tlForceNone, Unit:=unitNone, TName:="DIE_TYPE", lowVal:=0, hiVal:=20, formatStr:="%.0f")
'
'
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function
Public Function DecodeDeidFromOtp() '(b_print_IEDA As Boolean)
    Dim sFuncName As String: sFuncName = "DecodeDeidFromOtp"
    On Error GoTo ErrHandler
    Dim lChipIdIdx As Long
    Dim IIdx As Integer
    Dim sCh1st As String
    Dim sCh2to6 As String
    Dim iAscVal As Integer
    Dim sComment As String
    Dim iChkVal As Integer
    Dim sDeidBinStrM As String
    Dim lDecimal As Long
    Dim sBinStr As String
    Dim sLotID As String
    Dim sWaferID As String
    Dim sXcoordBinStr As String
    Dim sYcoordBinStr As String
    Dim sOtpRevBinStr As String
    Dim sTName As String
    Dim svDeidBinStrM As New SiteVariant
    Dim lOtpIdx As Long
    Dim sDEIDActStrM As String
    
    
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "----------------------------------------------------"
    TheExec.Datalog.WriteComment sFuncName + ":"
    
    For Each Site In TheExec.Sites
        TheExec.Datalog.WriteComment "----------------------------------------------------"
        If g_sbOtpedECID(Site) = True Then
            ''''per 8 OTP_Reg() to convert to binary string
            
            '20190725 Reset, or the value is the last one
            g_svOTP_LotID = ""
            sDeidBinStrM = ""
            For lChipIdIdx = 0 To UBound(g_asChipIDName)
                lOtpIdx = SearchOtpIdxByName(g_asChipIDName(lChipIdIdx))
                lDecimal = g_OTPData.Category(lOtpIdx).Read.Value(Site)
                sBinStr = g_OTPData.Category(lOtpIdx).Read.BitStrM(Site)
                If sBinStr = Empty Then sBinStr = String(g_OTPData.Category(lOtpIdx).lBitWidth, "0")
                sDeidBinStrM = sDeidBinStrM + StrReverse(sBinStr)
                TheExec.Datalog.WriteComment "Site(" & Site & ") Reg_Read(" & lChipIdIdx & ") = " & FormatLog(lDecimal, 4) & " [" + sBinStr + "]" & " , &H" & Right("00" & Hex(lDecimal), 2)
            Next lChipIdIdx
            
            'KC ECID Full 64bit string for control+F to work
            'TheExec.Datalog.WriteComment "Site(" & Site & ") ECID := " & "&H " & Right("00" & Hex(Reg_read(0)(Site)), 2) & Right("00" & Hex(Reg_read(1)(Site)), 2) & Right("00" & Hex(Reg_read(2)(Site)), 2) & Right("00" & Hex(Reg_read(3)(Site)), 2) & Right("00" & Hex(Reg_read(4)(Site)), 2) & Right("00" & Hex(Reg_read(5)(Site)), 2) & Right("00" & Hex(Reg_read(6)(Site)), 2) & Right("00" & Hex(Reg_read(7)(Site)), 2)
            
            TheExec.Datalog.WriteComment "Site(" & Site & ") Read DEID BinStrM = " + sDeidBinStrM + " [bit0...bit63][MSB...LSB]"
            TheExec.Datalog.WriteComment "----------------------------------------------------"
            
  
            sLotID = Mid(sDeidBinStrM, g_iLOTID_BITS_START + 1, g_iLOTID_BITS_BW)
            sWaferID = Mid(sDeidBinStrM, g_iWFID_BITS_START + 1, g_iWFID_BITS_BW)
            sXcoordBinStr = Mid(sDeidBinStrM, g_iXCOORD_BITS_START + 1, g_iXCOORD_BITS_BW)
            sYcoordBinStr = Mid(sDeidBinStrM, g_iYCOORD_BITS_START + 1, g_iYCOORD_BITS_BW)
            sOtpRevBinStr = Mid(sDeidBinStrM, g_iOTPREV_BITS_START + 1, g_iOTPREV_BITS_BW)
            
            
            For IIdx = 1 To 6
                g_svOTP_LotID = g_svOTP_LotID + ConvertLotIdBin2Letter(Mid(sLotID, 1 + (IIdx - 1) * 6, 6))
            Next IIdx
            
            'g_svOTP_LotID = CLng(ConvertFormat_Bin2Dec(sLotID))
            g_slOTP_WaferId = CLng(ConvertFormat_Bin2Dec(sWaferID))
            g_slOTP_XCoord = CLng(ConvertFormat_Bin2Dec(sXcoordBinStr))
            g_slOTP_YCoord = CLng(ConvertFormat_Bin2Dec(sYcoordBinStr))
            g_slOTP_Rev = CLng(ConvertFormat_Bin2Dec(sOtpRevBinStr))
            svDeidBinStrM = sDeidBinStrM
            TheExec.Datalog.WriteComment "Site(" & Site & ") OTP Read LotID   = " + CStr(g_svOTP_LotID)
            TheExec.Datalog.WriteComment "Site(" & Site & ") OTP Read WaferID = " + CStr(g_slOTP_WaferId)
            TheExec.Datalog.WriteComment "Site(" & Site & ") OTP Read X_Coord = " + CStr(g_slOTP_XCoord)
            TheExec.Datalog.WriteComment "Site(" & Site & ") OTP Read Y_Coord = " + CStr(g_slOTP_YCoord)
            'TheExec.Datalog.WriteComment "Site(" & Site & ") OTP Read OTP_Rev = " + CStr(g_slOTP_Rev)
            TheExec.Datalog.WriteComment "----------------------------------------------------"
            TheExec.Datalog.WriteComment ""


            sDEIDActStrM = g_svOTP_LotID + Format(g_slOTP_WaferId, "00") + Format(g_slOTP_XCoord(Site), "00") + Format(g_slOTP_YCoord(Site), "00")                                             '20170817
            TheExec.Datalog.WriteComment "Site(" & Site & ") Read DEID ActStrM = " + sDEIDActStrM + " [ LotID + WaferID + XCoord + YCoord ] "
            Call TheExec.Datalog.Setup.SetPRRPartInfo(tl_SelectSite, Site, , sDEIDActStrM)
        
        Else ' g_sbOtpedECID(Site) = False
            TheExec.Datalog.WriteComment "----------------------------------------------------"
            TheExec.Datalog.WriteComment "Site(" & Site & ") is a Fresh die (No ECID from OTP)"
            sDEIDActStrM = g_sLotID + Format(g_lWaferID, "00") + Format(g_slXCoord(Site), "00") + Format(g_slYCoord(Site), "00")                                             '20170817
            TheExec.Datalog.WriteComment "Site(" & Site & ") SetXY ActStrM = " + sDEIDActStrM + " [ LotID + WaferID + XCoord + YCoord ] "
        End If
        '''-----------------------------------------------------------------------------
    Next Site
    
    For Each Site In TheExec.Sites
        If g_sbOtpedECID(Site) = True Then
            TheExec.Datalog.WriteComment ("Prober_" + UCase(g_sLotID) + "_vs_DUT_" + UCase(g_svOTP_LotID))
            sTName = "Prober_vs_DUT"
            If (UCase(g_sLotID) = UCase(g_svOTP_LotID)) Then
                ''''Pass
                TheExec.Flow.TestLimit 1, 1, 1, TName:=sTName, formatStr:="%.0f", TNum:=9002000
            Else
                ''''Fail
                If TheExec.Datalog.Setup.LotSetup.TESTMODE = 3 Or g_bEnWrdOTPFTProg = True Then
                   TheExec.Flow.TestLimit 0, 1, 1, TName:=sTName, formatStr:="%.0f", ForceResults:=tlForcePass, TNum:=9002000
                Else
                    TheExec.Flow.TestLimit 0, 1, 1, TName:=sTName, formatStr:="%.0f", TNum:=9002000
                End If
            End If
            '3 = E = Engineering mode:ForceFlowPass
            If (TheExec.Datalog.Setup.LotSetup.TESTMODE = 3 Or g_bEnWrdOTPFTProg = True) Then
                TheExec.Flow.TestLimit g_lWaferID, g_slOTP_WaferId, g_slOTP_WaferId, TName:="DUT_Wafer_ID", formatStr:="%.0f", ForceResults:=tlForcePass, TNum:=9002001
                TheExec.Flow.TestLimit g_slXCoord, g_slOTP_XCoord, g_slOTP_XCoord, TName:="DUT_X_Coord", formatStr:="%.0f", ForceResults:=tlForcePass, TNum:=9002002
                TheExec.Flow.TestLimit g_slYCoord, g_slOTP_YCoord, g_slOTP_YCoord, TName:="DUT_Y_Coord", formatStr:="%.0f", ForceResults:=tlForcePass, TNum:=9002003
            Else
                TheExec.Flow.TestLimit g_lWaferID, g_slOTP_WaferId, g_slOTP_WaferId, TName:="DUT_Wafer_ID", formatStr:="%.0f", TNum:=9002001
                TheExec.Flow.TestLimit g_slXCoord, g_slOTP_XCoord, g_slOTP_XCoord, TName:="DUT_X_Coord", formatStr:="%.0f", TNum:=9002002
                TheExec.Flow.TestLimit g_slYCoord, g_slOTP_YCoord, g_slOTP_YCoord, TName:="DUT_Y_Coord", formatStr:="%.0f", TNum:=9002003
            End If
        End If
    Next Site

'    Else 'TheExec.CurrentJob NOT Like "*CP*"
    
    For Each Site In TheExec.Sites
        If g_sbOtpedECID(Site) = True Then 'And g_sbOtpedPGM(Site) = True
            ''''Syntax Check LotID of the DUT from Read OTD
            sCh1st = Mid(g_svOTP_LotID, 1, 1)
            iAscVal = Asc(LCase(sCh1st))
            If (iAscVal < 97 Or iAscVal > 122) Then ''''a=97 and z=122 in ANSI character
                iChkVal = 0 'Fail
                sComment = "First Character of DUT LotID (" + UCase(sCh1st) + ") is not [A-Z]."
                TheExec.Datalog.WriteComment sComment
            Else
                iChkVal = 1 'Pass
                If (Len(g_svOTP_LotID) <> 6) Then
                    sComment = "Character Numbers of Prober LotID (" + UCase(g_svOTP_LotID) + ") is NOT Six Characters."
                    TheExec.Datalog.WriteComment sComment
                    iChkVal = 0 'Fail
                Else
                    For IIdx = 2 To 6
                        sCh2to6 = Mid(g_svOTP_LotID, IIdx, 1)
                        iAscVal = Asc(LCase(sCh2to6))
                        If iAscVal < 97 Or iAscVal > 122 Then    'a=97 and z=122 in ANSI character
                            If iAscVal < 48 Or iAscVal > 57 Then ''0'=48 and '9'=57 in ANSI character
                                iChkVal = 0  'Fail
                                sComment = "Second-to-Sixth Characters of DUT LotID (" + UCase(g_svOTP_LotID) + ") are not [A-Z] or [0-9]."
                                TheExec.Datalog.WriteComment sComment
                                Exit For
                            Else
                                iChkVal = 1 'Pass
                            End If
                        Else
                        End If
                    Next IIdx
                End If
            End If
            sTName = "DUT_LotID_" + UCase(g_svOTP_LotID)
            TheExec.Flow.TestLimit iChkVal, 1, 1, TName:=sTName
            'TheExec.Flow.TestLimit iChkVal, -999,  999, TName:=tsname
            TheExec.Flow.TestLimit g_slOTP_WaferId, 1, 25, TName:="DUT_Wafer_ID", formatStr:="%.0f"
            TheExec.Flow.TestLimit g_slOTP_XCoord, g_iXCOORD_LOW_LMT, g_iXCOORD_HI_LMT, TName:="DUT_X_Coord", formatStr:="%.0f"
            TheExec.Flow.TestLimit g_slOTP_YCoord, g_iYCOORD_LOW_LMT, g_iYCOORD_HI_LMT, TName:="DUT_Y_Coord", formatStr:="%.0f"
        End If
    Next Site

    For Each Site In TheExec.Sites
        If g_sbOtpedECID(Site) = True And g_sbOtpedPGM(Site) = True Then
            sDEIDActStrM = UCase(g_svOTP_LotID) + CStr(Format(g_slOTP_WaferId(Site), "00")) + CStr(Format(g_slOTP_XCoord(Site), "00")) + CStr(Format(g_slOTP_YCoord(Site), "00"))
            TheExec.Datalog.WriteComment "Site(" & Site & ") READ DUT ECID READBACK ActStrM = " + sDEIDActStrM + " [ LotID + WaferID + XCoord + YCoord ] "
            Call TheExec.Datalog.Setup.SetPRRPartInfo(tl_SelectSite, Site, , sDEIDActStrM)
        End If
    Next Site

'    End If 'END IF TheExec.CurrentJob Like "*CP*"


    If UCase(TheExec.DataManager.InstanceName) Like "*LOCK*" Then
        For Each Site In TheExec.Sites.Existing
            If g_sbOtpedECID(Site) = True Then 'And g_sbOtpedPGM(Site) = True
                'ECID Burned die --> SetXY From OTP
                'If (UCase(g_sLotID) <> UCase(g_svOTP_LotID)) Then
                    '0 = A = AEL (Automatic Edge Lock) mode
                    '1 = C = Checker mode
                    '2 = D = Development/Debug test mode
                    '3 = E = Engineering mode (same as Development mode)
                    '4 = M = Maintenance mode
                    '5 = P = Production test mode
                    '6 = Q = Quality Control
                    If TheExec.Datalog.Setup.LotSetup.TESTMODE = 3 Or g_bEnWrdOTPFTProg = True Then
                        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(Site, g_slOTP_XCoord(Site))
                        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(Site, g_slOTP_YCoord(Site))
                    End If
                'End If
            End If
        Next Site
    End If
  ''''Add register and print out the IEDA data
  ''''register and print out the IEDA data----------------------------------------------------------------------------
    'If (b_print_IEDA) Then
    If UCase(TheExec.DataManager.InstanceName) Like "*READ*" Then
        Dim sDeidLotID As String: sDeidLotID = ""
        Dim sDeidWfID As String: sDeidWfID = ""
        Dim sDeidXCoord As String: sDeidXCoord = ""
        Dim sDeidYCoord As String: sDeidYCoord = ""
        Dim sDeidDeid As String: sDeidDeid = ""

        For Each Site In TheExec.Sites.Existing
            If g_sbOtpedECID(Site) = True Then 'And g_sbOtpedPGM(Site) = True
                If (Site = TheExec.Sites.Existing.Count - 1) Then
                    sDeidLotID = sDeidLotID + g_svOTP_LotID(Site)
                    sDeidWfID = sDeidWfID + CStr(g_slOTP_WaferId(Site))
                    sDeidXCoord = sDeidXCoord + CStr(g_slOTP_XCoord(Site))
                    sDeidYCoord = sDeidYCoord + CStr(g_slOTP_YCoord(Site))
                    sDeidDeid = sDeidDeid + svDeidBinStrM(Site)
                Else
                    sDeidLotID = sDeidLotID + g_svOTP_LotID(Site) + ","
                    sDeidWfID = sDeidWfID + CStr(g_slOTP_WaferId(Site)) + ","
                    sDeidXCoord = sDeidXCoord + CStr(g_slOTP_XCoord(Site)) + ","
                    sDeidYCoord = sDeidYCoord + CStr(g_slOTP_YCoord(Site)) + ","
                    sDeidDeid = sDeidDeid + svDeidBinStrM(Site) + ","
                End If
    
                'SetXY From OTP
                If TheExec.Sites(Site).Active = True Then
                    'If (UCase(g_sLotID) <> UCase(g_svOTP_LotID)) Then
                        If TheExec.Datalog.Setup.LotSetup.TESTMODE = 3 Or g_bEnWrdOTPFTProg = True Then
                            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(Site, g_slOTP_XCoord(Site))
                            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(Site, g_slOTP_YCoord(Site))
                        End If
                    'End If
                    
                    Dim device_code As String
                    device_code = g_svOTP_LotID(Site) & "_W" & CStr(g_slOTP_WaferId(Site)) & "_X" _
                                    & CStr(g_slOTP_XCoord(Site)) & "_Y" & CStr(g_slOTP_YCoord(Site)) & "_S" & Site
    
                    TheExec.Datalog.WriteComment "DEVICE_CODE: " + device_code
                    
    
                    'KC ECID 05 30 2018
                    device_code = g_svOTP_LotID(Site) & "" & CStr(Format(g_slOTP_WaferId(Site), "00")) & "" _
                                    & CStr(Format(g_slOTP_XCoord(Site), "00")) & "" & CStr(Format(g_slOTP_YCoord(Site), "00"))
    
                    TheExec.Datalog.WriteComment "Site(" & Site & ") Write DEID ActStrM = " + device_code + " [ LotID + WaferID + XCoord + YCoord ] "
                    Call TheExec.Datalog.Setup.SetPRRPartInfo(tl_SelectSite, Site, , device_code)
                    TheExec.Datalog.WriteComment ""
                End If
            Else
                If (Site = TheExec.Sites.Existing.Count - 1) Then
                Else
                    sDeidLotID = sDeidLotID + ","
                    sDeidWfID = sDeidWfID + ","
                    sDeidXCoord = sDeidXCoord + ","
                    sDeidYCoord = sDeidYCoord + ","
                    sDeidDeid = sDeidDeid + ","
                End If
            End If
        Next Site

        sDeidLotID = CheckNCombineIEDA(sDeidLotID)
        sDeidWfID = CheckNCombineIEDA(sDeidWfID)
        sDeidXCoord = CheckNCombineIEDA(sDeidXCoord)
        sDeidYCoord = CheckNCombineIEDA(sDeidYCoord)
        sDeidDeid = CheckNCombineIEDA(sDeidDeid) '20170926

        TheExec.Datalog.WriteComment vbCrLf & "Test Instance : " + TheExec.DataManager.InstanceName
        TheExec.Datalog.WriteComment " ECID (all sites iEDA format): RegKey Path = " & "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\"
        TheExec.Datalog.WriteComment " Info Read Back FROM OTP"
        TheExec.Datalog.WriteComment " Lot ID    = " + sDeidLotID
        TheExec.Datalog.WriteComment " Wafer ID  = " + sDeidWfID
        TheExec.Datalog.WriteComment " X_Coor    = " + sDeidXCoord
        TheExec.Datalog.WriteComment " Y_Coor    = " + sDeidYCoord
        TheExec.Datalog.WriteComment " ECID_DEID = " + sDeidDeid & vbCrLf

        '============================================
        '=  Write Data to Register Edit (HKEY)      =
        '============================================
        Call RegKeySave("eFuseLotNumber", sDeidLotID)
        Call RegKeySave("eFuseWaferID", sDeidWfID)
        Call RegKeySave("eFuseDieX", sDeidXCoord)
        Call RegKeySave("eFuseDieY", sDeidYCoord)
        Call RegKeySave("HraECID_64bit", sDeidDeid)

    End If
    ''''End of register and print out the IEDA data----------------------------------------------------------------------------
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function BurnOTP(r_eSelect_Block As g_eOTPBLOCK_TYPE, r_psWriteSetupPat As PatternSet, r_psWritePat As PatternSet, _
                        r_psFWTsuDNGS As PatternSet, r_psFWTsuSYS As PatternSet, r_psFWTsuCRC As PatternSet, _
                        Optional Validating_ As Boolean)
    Dim sFuncName As String: sFuncName = "BurnOTP"
    On Error GoTo ErrHandler
    If Validating_ Then
        If r_psWriteSetupPat <> "" Then Call PrLoadPattern(r_psWriteSetupPat.Value)
        If r_psWritePat <> "" Then Call PrLoadPattern(r_psWritePat.Value)
        If r_psFWTsuDNGS <> "" Then Call PrLoadPattern(r_psFWTsuDNGS.Value)
        If r_psFWTsuSYS <> "" Then Call PrLoadPattern(r_psFWTsuSYS.Value)
        If r_psFWTsuCRC <> "" Then Call PrLoadPattern(r_psFWTsuCRC.Value)
        Exit Function
    End If
    Dim dVDDVVal             As Double
    Dim lBurnAddrStart        As Long
    Dim lBurnAddrEnd          As Long
    Dim dReferenceTime       As Double
    Dim asPatArray() As String
    Dim lPatCnt      As Long
    Dim psOTPRead    As New PatternSet
    Dim psOneShotWritePat As New PatternSet
    Dim sbSaveSiteStatus As New SiteBoolean
    
    '20190930 To prevent the long pattern load time during OnProgramValidated
    If g_bOTPOneShot = True And r_eSelect_Block = eECID_OTPBURN Then
        psOneShotWritePat.Value = g_sOTP_ONESHOT_WRITE
    End If

    '___Get pattern name from PatternSet
    If g_bOTPFW = False Then 'FWDebug
        r_psWriteSetupPat.Value = GetPatListFromPatternSet_OTP(r_psWriteSetupPat.Value, asPatArray, lPatCnt)
        r_psWritePat.Value = GetPatListFromPatternSet_OTP(r_psWritePat.Value, asPatArray, lPatCnt)
        psOneShotWritePat.Value = GetPatListFromPatternSet_OTP(psOneShotWritePat.Value, asPatArray, lPatCnt)
    Else
        '___20200313 add, need to check
        r_psFWTsuDNGS.Value = GetPatListFromPatternSet_OTP(r_psFWTsuDNGS.Value, asPatArray, lPatCnt)
        r_psFWTsuSYS.Value = GetPatListFromPatternSet_OTP(r_psFWTsuSYS.Value, asPatArray, lPatCnt)
        r_psFWTsuCRC.Value = GetPatListFromPatternSet_OTP(r_psFWTsuCRC.Value, asPatArray, lPatCnt)
    End If

     
    Call ProfileMark("OTP_Burn:T0")
    sbSaveSiteStatus = TheExec.Sites.Selected
    If TheExec.Sites.Selected.Count = 0 Then
       TheExec.Datalog.WriteComment "*** No Site is alive; bypass " & sFuncName & " **************************************************"
       Exit Function
    Else
        '20190704 Print Selected BurnOTP_Type
        If (g_bOTPFW = False) And (g_bOTPOneShot = False) Then TheExec.Datalog.WriteComment "**********************OTP Multi-Shot Burn is executing!!!**********************"
        If g_bOTPFW = True Then TheExec.Datalog.WriteComment "**********************OTP FW Burn is executing!!!**********************"
        If g_bOTPOneShot = True Then TheExec.Datalog.WriteComment "**********************OTP OneShot Burn is executing!!!**********************"
    End If

    '___OTP Burn Addr Range Definition
    Select Case r_eSelect_Block
        Case (eECID_OTPBURN)
            TheExec.Datalog.WriteComment "<" + sFuncName + "===>(ECID_OTPBURN)" + ">"
            If g_sOTPRevisionType = "OTP_V01" Then 'V1 Burn ECID only
                lBurnAddrStart = 0
                lBurnAddrEnd = 2
            Else 'other versions burn all data except CRC will be burned after CRC calculation
                lBurnAddrStart = 0
                lBurnAddrEnd = g_iOTP_ADDR_END - 1 '1023
            End If
        Case (eCRC_OTPBURN)
            TheExec.Datalog.WriteComment "<" + sFuncName + "===>(CRC_OTPBURN)" + ">"
            If g_sOTPRevisionType = "OTP_V01" Or g_bOTPFW = True Then
                Exit Function
            Else 'other versions burn CRC after CRC calculation
                lBurnAddrStart = g_iOTP_ADDR_END '1024
                lBurnAddrEnd = g_iOTP_ADDR_END
            End If
    End Select
   
    dReferenceTime = TheExec.Timer
      
    TheExec.Datalog.WriteComment ("*** OTP_updatetrimcodes Time: " & TheExec.Timer(dReferenceTime) & " ***")
    '___Disconnect all digital pins
    TheHdw.PPMU.Pins(g_sDIG_PINS).Disconnect

    Call ProfileMark("BurnOTP:T1:")

    dVDDVVal = TheHdw.DCVI.Pins(g_sVDD_PINNAME).Voltage
 
    If TheExec.Sites.Selected.Count = 0 Then GoTo skip_programming
 
    '___[[OTP_Brun_Write]]
    Call BurnWriteOTP(r_eSelect_Block, r_psWriteSetupPat, r_psWritePat, psOneShotWritePat, r_psFWTsuDNGS, r_psFWTsuSYS, r_psFWTsuCRC, lBurnAddrStart, lBurnAddrEnd)
    '___[[OTP_Brun_Read]]
skip_programming:
     
TheExec.Sites.Selected = sbSaveSiteStatus

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function ReadOTP(r_eSelect_Block As g_eOTPBLOCK_TYPE, r_psReadSetupPat As PatternSet, r_psReadPat As PatternSet, _
                         Optional Validating_ As Boolean)
    Dim sFuncName As String: sFuncName = "ReadOTP"
    On Error GoTo ErrHandler
    If Validating_ Then
        If r_psReadSetupPat <> "" Then Call PrLoadPattern(r_psReadSetupPat.Value)
        If r_psReadPat <> "" Then Call PrLoadPattern(r_psReadPat.Value)
        Exit Function
    End If
    Dim dVDDVVal             As Double
    Dim aslOtpChipIDReadVal()  As New SiteLong
    ReDim aslOtpChipIDReadVal(UBound(g_asChipIDName))
    Dim lActAddr             As Long
    Dim lReadAddrStart        As Long
    Dim lReadAddrEnd          As Long
    Dim sbSaveSiteStatus As New SiteBoolean
    Dim asPatArray() As String
    Dim lPatCnt As Long
    Dim lChipIdIdx As Long
    Dim lOtpIdx As Long
    Dim psOneShotReadPat As New PatternSet
    
    ''''20190723 new method by Ching
    Dim wPGMData As New DSPWave  ''''it's the decimal of 32-bits(gD_slOTP_REGDATA_BW)
    Dim wReadData As New DSPWave ''''it's the decimal of 32-bits(gD_slOTP_REGDATA_BW)
    Dim wPGMReadDiff As New DSPWave
    ''''Site boolean to do the site control
    Dim sbOldBurnedDie As New SiteBoolean
    Dim sbNewBurnedDie As New SiteBoolean
    Dim sbKeepFreshDie As New SiteBoolean


If g_bOTPFW = True And g_bFWDlogCheck = False Then Exit Function

'If gFlag_POPEnd = False Then Call Setup_JTAG_nWireSPMI  'Jeff added this in case IDAC POP_end was not executed. 11/05/19

    '___OTP Burn Addr Range Definition
'    If g_sOTPRevisionType = "OTP_V01" Then '=V1 Burn ECID only
'       lReadAddrStart = 0
'       lReadAddrEnd = 2
'    Else
        lReadAddrStart = 0
        lReadAddrEnd = g_iOTP_ADDR_END
'    End If
    
    '20190930 To prevent the long pattern load time during OnProgramValidated
    If g_bOTPOneShot = True Then
        psOneShotReadPat.Value = g_sOTP_ONESHOT_READ
    End If

    '___Get pattern name from PatternSet
    If g_bOTPFW = False Then 'FWDebug
        r_psReadSetupPat.Value = GetPatListFromPatternSet_OTP(r_psReadSetupPat.Value, asPatArray, lPatCnt)
        r_psReadPat.Value = GetPatListFromPatternSet_OTP(r_psReadPat.Value, asPatArray, lPatCnt)
        If g_bOTPOneShot = True Then psOneShotReadPat.Value = GetPatListFromPatternSet_OTP(psOneShotReadPat.Value, asPatArray, lPatCnt)
     Else
        '___20200313 add, need to check
        r_psReadPat.Value = GetPatListFromPatternSet_OTP(r_psReadPat.Value, asPatArray, lPatCnt)
        If g_bOTPOneShot = True Then psOneShotReadPat.Value = GetPatListFromPatternSet_OTP(psOneShotReadPat.Value, asPatArray, lPatCnt)
     End If
     
    Call ProfileMark("ReadOTP:T0")
    sbSaveSiteStatus = TheExec.Sites.Selected
    
    For Each Site In TheExec.Sites
        If TheExec.Sites(Site).SiteVariableValue("RunTrim") <> -1 Then  'Burned -> Burned 'And g_bOtpEnable = True
            TheExec.Datalog.WriteComment "Site (" & Site & ") is OldBurnedDie"
            sbOldBurnedDie(Site) = True
        ElseIf TheExec.Sites(Site).SiteVariableValue("RunTrim") = -1 And g_bOtpEnable = True Then 'Fresh -> Burned
            TheExec.Datalog.WriteComment "Site (" & Site & ") is NewBurnedDie (either Trim/ECID)"
            sbNewBurnedDie(Site) = True
        ElseIf TheExec.Sites(Site).SiteVariableValue("RunTrim") = -1 And g_bOtpEnable = False Then 'Fresh -> Fresh
            TheExec.Datalog.WriteComment "Site (" & Site & ") is FreshDie"
            sbKeepFreshDie(Site) = True
        End If
    Next Site
    
    'If TheExec.Sites.Selected.Count = 0 Then
    If sbKeepFreshDie.All(True) Then
       TheExec.Datalog.WriteComment "*** All sites are Fresh; bypass " & sFuncName & " **************************************************"
       TheExec.Sites.Selected = sbSaveSiteStatus
       Exit Function
    Else
        '20190704 Print Selected OTPRead_Type
        If g_bOtpEnable = True And (g_bOTPFW = False) And (g_bOTPOneShot = False) Then TheExec.Datalog.WriteComment "**********************OTP Multi-Shot Read is executing!!!**********************"
        'If g_bOTPFW = True Then TheExec.Datalog.WriteComment "**********************OTP FW Burn is executing!!!**********************"
        If g_bOTPOneShot = True Then TheExec.Datalog.WriteComment "**********************OTP OneShot Read is executing!!!**********************"
    End If
    

    '___Disconnect all digital pins
    TheHdw.PPMU.Pins(g_sDIG_PINS).Disconnect

    dVDDVVal = TheHdw.DCVI.Pins(g_sVDD_PINNAME).Voltage
    '___[[OTP_Brun_Print_Datalog]]
    m_VddLevel = "_VDD" & CStr(Format(dVDDVVal, "0.00"))
    
    
    TheHdw.DCVI.Pins("ATB_DC30_ALL").Current = 20 * uA  ' Added for avoiding the open loop alarm 1004/2019
    If g_bOTPOneShot = True Or g_bFWDlogCheck = True Then
''    TheHdw.DCVI.Pins("ATB_DC30_ALL").Current = 20 * uA  ' Added for avoiding the open loop alarm 1004/2019
        Call OTP_READREG_DSP_ALL(psOneShotReadPat) 'One shot
    ElseIf g_bOTPFW = False And g_bOTPOneShot = False Then
        If (True) Then
            'Dim Exetime As Double
            For lActAddr = lReadAddrStart To lReadAddrEnd
               Call OTP_READREG_DSP(r_psReadPat, lActAddr) 'Multi-shot
            Next lActAddr
            'Exetime = TheHdw.ReadStopwatch
            'TheExec.AddOutput ("Outside ForLoop" & Exetime)
        Else 'Need online verification
            'TheHdw.StartStopwatch
            Call OTP_READREG_DSP_LoopAddr(r_psReadPat.Value, lReadAddrStart, lReadAddrEnd)
            'Exetime = TheHdw.ReadStopwatch
            'TheExec.AddOutput ("Inside ForLoop" & Exetime)
        End If
    End If
    
    If g_sOTPRevisionType = "OTP_V01" Then
        For lChipIdIdx = 0 To UBound(g_asChipIDName)
            lOtpIdx = g_OTPData.Category(SearchOtpIdxByName(g_asChipIDName(lChipIdIdx))).lOtpIdx
            Call SetReadData2OTPCat_byOTPIdx(lOtpIdx)
        Next lChipIdIdx
    Else
        For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
            Call SetReadData2OTPCat_byOTPIdx(lOtpIdx)
        Next lOtpIdx
    End If

        '20200401 JY Fuji would like to double confirm 'only ECID' are burnt
    If g_sOTPRevisionType = "OTP_V01" Then
        Dim lEcidBitEnd As Long: lEcidBitEnd = g_OTPData.Category(SearchOtpIdxByName(g_asChipIDName(UBound(g_asChipIDName)))).lOtpBitStrEnd
        Dim slSumUpRead As New SiteLong
    
        Call RunDsp.SumUpCatReadData_ExceptECID(lEcidBitEnd + 1, CLng(g_iOTP_ADDR_TOTAL) * g_iOTP_REGDATA_BW - lEcidBitEnd, slSumUpRead)
        Call TheExec.Flow.TestLimit(TName:="SUM_UP_READ_VALUE_EXCEPT_ECID", ResultVal:=slSumUpRead, lowVal:=0, hiVal:=0, formatStr:="%.0f", scaletype:=scaleNoScaling) ' formatStr:="%.6f") '
        GoTo ResetSiteStatus
        
    End If

OldBurnedDie:
    'Old burned die didn't go through trim section, it's meaningless to compare gD_wPGMData / gD_wReadData
    TheExec.Sites.Selected = sbOldBurnedDie
    If TheExec.Sites.Selected.Count = 0 Then GoTo NewBurnedDie
    Call RunDsp.otp_Read_DataWave(wReadData)
    If g_bExpectedActualDebugPrint Then Call CompareExpActDatalog_read_only(lReadAddrStart, lReadAddrEnd, wReadData)
    
NewBurnedDie:
    'New burned die need to compare gD_wReadData are same as what(gD_wPGMData) we expected to burn into OTP
    TheExec.Sites.Selected = sbNewBurnedDie
    If TheExec.Sites.Selected.Count = 0 Then GoTo ResetSiteStatus
    If g_bOTPFW = True Then
        '___Decode ECID with AHBRead
        For lChipIdIdx = 0 To UBound(g_asChipIDName)
                g_RegVal = aslOtpChipIDReadVal(lChipIdIdx)
                Call auto_OTPCategory_GetReadDecimal_AHB(CStr(g_asChipIDName(lChipIdIdx)), g_RegVal) ', False)
        Next lChipIdIdx
        
        '___Set OTP read data
        If TheExec.TesterMode = testModeOffline Then
            For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
                g_OTPData.Category(lOtpIdx).Read.Value = g_OTPData.Category(lOtpIdx).Write.Value
            Next lOtpIdx
        Else
            '20190701
            Dim lStartPnt As Long
            Dim lEndtPnt As Long
            For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
                lStartPnt = g_OTPData.Category(lOtpIdx).lOtpBitStrStart
                lEndtPnt = g_OTPData.Category(lOtpIdx).lOtpBitStrEnd
                g_OTPData.Category(lOtpIdx).Read.Value = g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs
                '20190701 Add binary stream for DEID Decode check
                For Each Site In TheExec.Sites
                    g_OTPData.Category(lOtpIdx).Read.HexStr(Site) = "0x" + Hex(g_OTPData.Category(lOtpIdx).Read.Value(Site))
                    g_OTPData.Category(lOtpIdx).Read.BitStrM(Site) = ConvertFormat_Dec2Bin_Complement(g_OTPData.Category(lOtpIdx).Read.Value(Site), lEndtPnt - lStartPnt + 1)
                Next Site
            Next lOtpIdx
        End If
        
        If g_bFWDlogCheck = True Then
            Call RunDsp.otp_compare_PGM_Read_DataWave(wPGMData, wReadData, wPGMReadDiff)
            If g_bExpectedActualDebugPrint Then Call CompareExpActDatalog(lReadAddrStart, lReadAddrEnd, wPGMData, wReadData, wPGMReadDiff)
        End If
    Else
        '20190107 The reason why we need to setWrite here is because ECID burned die will fail in Expect/Actual compare
        '----------------------------------------------------------------------------------------------------------------
        Dim wChipIdReg As New DSPWave
        Dim lIdx As Long
        Dim alTemp() As Long
        Dim lDecimal As Long
        Dim sBinStr As String
        Dim sDeidBinStrM As String: sDeidBinStrM = ""
        For lChipIdIdx = 0 To UBound(g_asChipIDName)
            lOtpIdx = SearchOtpIdxByName(g_asChipIDName(lChipIdIdx))
            lDecimal = g_OTPData.Category(lOtpIdx).Read.Value(Site)
            sBinStr = g_OTPData.Category(lOtpIdx).Read.BitStrM(Site)
            If sBinStr = Empty Then sBinStr = String(g_OTPData.Category(lOtpIdx).lBitWidth, "0")
            sDeidBinStrM = sDeidBinStrM + StrReverse(sBinStr)
            'THEEXEC.Datalog.WriteComment "Site(" & Site & ") Reg_Read(" & lChipIdIdx & ") = " & FormatLog(lDecimal, 4) & " [" + sBinStr + "]" & " , &H" & Right("00" & Hex(lDecimal), 2)
        Next lChipIdIdx
        For Each Site In TheExec.Sites
            Call otp_allotECID(wChipIdReg)
            For lIdx = 0 To 7 Step 1
                lDecimal = wChipIdReg(Site).Element(lIdx)
                g_aslOTPChipReg(lIdx)(Site) = lDecimal
                sBinStr = ConvertFormat_Dec2Bin(lDecimal, 8, alTemp)
                'THEEXEC.Datalog.WriteComment "Site(" & Site & ") g_aslOTPChipReg(" & lIdx & ") = " & FormatLog(lDecimal, 4) & " [" + sBinStr + "]"
            Next lIdx
        Next Site
        For lChipIdIdx = 0 To UBound(g_asChipIDName)
            Call auto_OTPCategory_SetWriteDecimal(g_asChipIDName(lChipIdIdx), g_aslOTPChipReg(lChipIdIdx))
        Next lChipIdIdx
        '----------------------------------------------------------------------------------------------------------------
        '** Normal Case:
        '___Print the decimal results in each address
        Call RunDsp.otp_compare_PGM_Read_DataWave(wPGMData, wReadData, wPGMReadDiff)
        If g_bExpectedActualDebugPrint Then Call CompareExpActDatalog(lReadAddrStart, lReadAddrEnd, wPGMData, wReadData, wPGMReadDiff)
    End If

ResetSiteStatus:
    TheExec.Sites.Selected = sbSaveSiteStatus
    Call ProfileMark("OTPRead:T4:OTP Datalog")
    
    Call DecodeDeidFromOtp

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function BurnWriteOTP(r_eSelectBlock As g_eOTPBLOCK_TYPE, r_psWriteSetupPat As PatternSet, r_psWritePat As PatternSet, r_psOneShotWritePat As PatternSet, _
                                r_psFWTsuDNGS As PatternSet, r_psFWTsuSYS As PatternSet, r_psFWTsuCRC As PatternSet, _
                                r_lBurnAddrStart As Long, r_lBurnAddrEnd As Long)
    Dim sFuncName As String: sFuncName = "BurnWriteOTP"
    On Error GoTo ErrHandler
    Dim lBurnAddrIdx As Long
    Dim lActAddr As Long
    Dim wAddrAndData As New DSPWave
    '__for mask purpose
    Dim wMaskECID As New DSPWave
    Dim wAllZero As New DSPWave
    Dim wIndex As New DSPWave
    Dim lEcidBitStart As Long
    Dim lEcidBitEnd As Long
    Dim lSize As Long
    Dim lOtpIdx As Long

    '20190828 JY: For self test pattern
    '===============================================================================
    'Dim Selftestpat As New PatternSet
    'Dim FailCount As New PinListData
    'Dim l_FailCnt As New SiteLong
    'Dim b_SiteCtrl As New SiteBoolean
    
    
    '    AHB_SEL.Value = g_sAHB_SEL
    '    GetPatListNExecutePat AHB_SEL
        '___Force device status as to-be-burned under offline mode
    '    If (TheExec.TesterMode = testModeOffline) Then
    '        If Not g_sOTPRevisionType = "OTP_V01" Then
    '            For Each Site In TheExec.Sites
    '                'Since in offline lockbit check, force elementlite(0)=0
    '               'gD_wPGMData.ElementLite(0) = 1
    '               l_LCKBit_offline = 1
    '            Next Site
    '               Call auto_OTPCategory_SetWriteDecimal(g_sOTP_PRGM_BIT_REG, l_LCKBit_offline)
    '        End If
    '    End If
    '===============================================================================

    '___EFUSE POWER UP/JTG_EFUSE WRITE SETUP, Burst pattern: r_psWriteSetupPat
    Call OTP_WRITEREG_initPatt(r_psWriteSetupPat)
    
    '___OTP_PowerUp_Vpp
    If g_bVPP_DISABLE = True Then
    Else
    '___RAISING VPP to 7.5 V
        Call RampUpVpp_ForOtpBurn
    End If
    '20190828 JY: Used to program OTP debug write test setup1 - VPP =7.5V  at the beginning
    '===============================================================================
    
    ''''    Selftestpat.Value = g_sOTP_WRITE_SELF_TEST_PAT
    ''''    TheHdw.Patterns(Selftestpat).Load
    ''''    TheHdw.Patterns(Selftestpat).test pfAlways, 0
    ''''    TheHdw.Digital.Patgen.HaltWait
    ''''    FailCount = TheHdw.Digital.Pins("ACTIVE_READY").FailCount ' Get the fail count.
    ''''    l_FailCnt = FailCount.Pins("ACTIVE_READY")
    ''''    For Each Site In TheExec.Sites
    ''''        If l_FailCnt(Site) = 0 Then
    ''''            b_SiteCtrl = True
    ''''        Else
    ''''            b_SiteCtrl = False
    ''''            TheExec.Datalog.WriteComment ("Site" & Site & " FailCnt>0 -> Shut down !")
    ''''        End If
    ''''    Next Site
            'If self test pattern fail --> 1. Bin Out 2. Site off
    '        Call TheExec.Flow.TestLimit(l_FailCnt, lowVal:=0, hiVal:=0, TName:="BurnWriteOTP_Self_Test_Pattern_FailCnt", formatStr:="%.0f", ForceResults:=tlForceFlow)
    '        TheExec.Sites.Selected = b_SiteCtrl
    '===============================================================================
        
        
    '3. JTG_EFUSE WRITE SEQ_DSC
        TheExec.Datalog.WriteComment ("JTG_EFUSE WRITE SEQ_DSC")
        
        Select Case r_eSelectBlock
        
        Case (eECID_OTPBURN)
            If TheExec.TesterMode = testModeOffline Then
                If g_sOTPRevisionType = "OTP_V01" = False Then
                    For Each Site In TheExec.Sites
                       gD_wPGMData.Element(0) = 1
                    Next Site
                End If
            End If
            TheExec.Datalog.WriteComment ("RUN PAT:" & r_psWritePat.Value)
            '___[OTP Burn]FW
            '************
            '***FW OTP***
            '************
            If (g_bOTPFW = True) Then
                 Call BurnWriteOTP_FW(r_psFWTsuDNGS, r_psFWTsuSYS, r_psFWTsuCRC)
            '___[OTP Burn]General multi-shot or one-shot cases
            Else
                
                '___20200401 JY need to move before mask
                '___Before writing OTP, update the write value to DSPwave first.
                For lOtpIdx = 0 To UBound(g_alDefRealUpdate) - 1
                    For Each Site In TheExec.Sites
                        'If gDW_RealDef_fromWrite.ElementLite(lOTPIdx) = 1 Then 'means it ever ran through the setWriteDecimal update.
                        If g_alDefRealUpdate(lOtpIdx) = 1 Then
                            Call OTPData2DSPWave(lOtpIdx)
                        End If
                    Next Site
                Next lOtpIdx
                
                'for ECID mask ( if the chip has already ECID bruned )
                lEcidBitStart = g_OTPData.Category(SearchOtpIdxByName(g_asChipIDName(0))).lOtpBitStrStart
                lEcidBitEnd = g_OTPData.Category(SearchOtpIdxByName(g_asChipIDName(UBound(g_asChipIDName)))).lOtpBitStrEnd
                lSize = lEcidBitEnd - lEcidBitStart + 1
                wAllZero.CreateConstant 0, lSize, DspLong 'used as 'mask'
                wIndex.CreateRamp lEcidBitStart, 1, lSize, DspLong
                For Each Site In TheExec.Sites
                    wMaskECID = gD_wPGMData.Copy
                    Call wMaskECID.ReplaceElements(wIndex, wAllZero)
                Next Site
                

                '************
                '**One Shot**
                '************
                If g_bOTPOneShot = True Then
                    '___OTP-DSP 20190411
                    Call OTP_WRITEREG_DSP_ALL(r_psOneShotWritePat.Value, wMaskECID)
                Else
                '**************
                '**Multi Shot**
                '**************
                    'Dim Exetime As Double
                    Dim lOtpOfs As Long
                    'TheHdw.StartStopwatch
                    lOtpOfs = g_iOTP_ADDR_OFFSET

                    Call RunDsp.otp_get_pgm_AddrDataWave_maskECID_LoopAddr(g_sbOtpedECID, wMaskECID, lOtpOfs, _
                                                                           r_lBurnAddrStart, r_lBurnAddrEnd, wAddrAndData)
                    TheHdw.Wait 100 * us ''''check it later
                    For Each Site In TheExec.Sites
                        TheExec.Datalog.WriteComment vbTab & "Site(" & Site & ") AddrAndDataWave SampleSize = " & wAddrAndData.SampleSize
                    Next Site
                    Call OTP_WRITEREG_DSP_LoopAddr(r_psWritePat.Value, r_lBurnAddrStart, r_lBurnAddrEnd, wAddrAndData)
                    'Exetime = TheHdw.ReadStopwatch
                    'TheExec.AddOutput ("Inside ForLoop" & Exetime)
                End If
            End If
      
        Case (eCRC_OTPBURN)
            TheExec.Datalog.WriteComment ("RUN PAT:" & r_psWritePat.Value)
            For lBurnAddrIdx = r_lBurnAddrStart To r_lBurnAddrEnd
                '___OTP-DSP 20190411
                lActAddr = lBurnAddrIdx + g_iOTP_ADDR_OFFSET
                Call RunDsp.otp_get_pgm_AddrDataWave(lActAddr, lBurnAddrIdx, wAddrAndData)
                TheHdw.Wait 50 * us ''''check it later
                If TheExec.TesterMode = testModeOnline Then Call OTP_WRITEREG_DSP(r_psWritePat.Value, lBurnAddrIdx, wAddrAndData)
            Next lBurnAddrIdx
        End Select
               
    '___OTP_PowerDown_Vpp
    If g_bVPP_DISABLE = True Then
        'Stop
    Else
        Call RampDownVpp_ForOtpBurn
    End If
    
    
    Select Case r_eSelectBlock
        Case (eECID_OTPBURN)
            TheExec.Datalog.WriteComment "<" + sFuncName + "===>(OTP_OTPBURN)" + ">"

            If g_bOTPFW = False Then
                
                If g_sOTPRevisionType = "OTP_V01" Then 'V1 Burn ECID only
                    '___20200313, don't need SiteLoop 'Janet need check
                    g_sbOtpedECID = True
'                        For Each Site In TheExec.Sites
'                            g_sbOtpedECID(Site) = True
'                        Next Site
                Else 'other versions burn all data except CRC will be burned after CRC calculation
                    '___20200313, don't need SiteLoop
                    g_sbOtped = True
                    g_sbOtpedECID = True
                    g_sbOtpedPGM = True
'                        For Each Site In TheExec.Sites
'                            g_sbOtped(Site) = True
'                            g_sbOtpedECID(Site) = True
'                            g_sbOtpedPGM(Site) = True
'                        Next Site
                End If
            Else
                For Each Site In TheExec.Sites
                    g_sbOtped(Site) = True
                    g_sbOtpedECID(Site) = True
                    g_sbOtpedPGM(Site) = True
                    g_sbOtpedCRC(Site) = True
                Next Site
            End If
        Case (eCRC_OTPBURN)
            TheExec.Datalog.WriteComment "<" + sFuncName + "===>(CRC_OTPBURN)" + ">"
            If g_sOTPRevisionType = "OTP_V01" Or g_bOTPFW = True Then 'V1 ignores CRC
               'Exit Function
            Else 'other versions burn CRC after CRC calculation
              '___20200313, don't need SiteLoop
              g_sbOtpedCRC = True
'              For Each Site In TheExec.Sites
'                  g_sbOtpedCRC(Site) = True
'              Next Site
            End If
    End Select
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
                                                                                                        

Public Function CheckLCKBit(r_psAhbSelPat As PatternSet, r_psReadPat As PatternSet, _
                            Optional Validating_ As Boolean) As Long
    '___Check the Program bit and Lock bit
    '___Write ECID to AHB in the end of CheckLCKBit function
    Dim sFuncName As String: sFuncName = "CheckLCKBit"
    On Error GoTo ErrHandler
    If Validating_ Then
        Call PrLoadPattern(r_psAhbSelPat.Value)
        Call PrLoadPattern(r_psReadPat.Value)
        Exit Function
    End If
    Dim asPatArray()             As String
    Dim lPatCnt                  As Long
    Dim lActAddr                 As Long
    Dim lChipIdIdx               As Long
    Dim sDefaultValue            As String
    Dim slPGMBitValue            As New SiteLong
    Dim slECIDBitValue           As New SiteLong
    Dim slCRCBitValue            As New SiteLong
    'Dim mSL_PGM_CHECK_BIT         As New SiteLong
    Dim sLCKBitCheckTNameGrp5    As String
    'Dim sbSaveSiteStatus As New SiteBoolean
    Dim sTestName As String
    '*******************************
    ' OTP Test Name Conversion
    '*******************************
    ReDim Preserve g_asLogTestName(4)
    g_asLogTestName(0) = "OTP"
    g_asLogTestName(1) = "ECID"  '
    g_asLogTestName(2) = "CHECK"
    g_asLogTestName(3) = "PRE"
    g_asLogTestName(4) = ""
        
    ' **** EFUSE READ  Production------->  EFUSE POWER UP /JTG_EFUSE READ SEQ_DSC/*****************
    'TheExec.Datalog.WriteComment (r_psReadPat)
    TheHdw.PPMU.Pins(g_sDIG_PINS).Disconnect
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    GetPatListNExecutePat r_psAhbSelPat '20190329
    r_psReadPat.Value = GetPatListFromPatternSet_OTP(r_psReadPat.Value, asPatArray, lPatCnt)
    

    TheExec.Datalog.WriteComment "<" + sFuncName + ">"
    
    '___Step1. OTP_READREG_DSP to Check Burn or Not. Read LotID and CRC.
    '___If the device has been burned, then the program bit should be "1", and it is "0" vice versa.
    '___Read LotID instead of whole ECID, CRC is usually non-zero value if the device has been burned.
    
    '___Offline Simulation
    If TheExec.TesterMode = testModeOffline Then
        sDefaultValue = g_OTPData.Category(SearchOtpIdxByName(g_sOTP_PRGM_BIT_REG)).lDefaultValue
        If (UCase(TheExec.CurrentJob) = "QA") Or (UCase(TheExec.CurrentJob) Like "*CHAR*") Then
            g_OTPData.Category(SearchOtpIdxByName(g_sOTP_PRGM_BIT_REG)).lDefaultValue = sDefaultValue
        Else
            '___Force device status as non-otped under offline mode
            For Each Site In TheExec.Sites
               gD_wPGMData.ElementLite(0) = 0
            Next Site
            'g_OTPData.Category(SearchOtpIdxByName(g_sOTP_PRGM_BIT_REG)).lDefaultValue = 0
        End If

    Else
        For lActAddr = 0 To g_OTPData.Category(SearchOtpIdxByName(g_asChipIDName(UBound(g_asChipIDName)))).lOtpA0
            Call OTP_READREG_DSP(r_psReadPat.Value, lActAddr)
        Next lActAddr
        Call OTP_READREG_DSP(r_psReadPat.Value, g_iOTP_ADDR_END) ''''20200313, it's CRC
    End If
    
    '___Step2. Update the data.Set the readback data into g_OTPData.category.read
    Call SetReadData2OTPCat_byOTPIdx(SearchOtpIdxByName(g_sOTP_PRGM_BIT_REG))
    Call SetReadData2OTPCat_byOTPIdx(SearchOtpIdxByName(g_sOTP_CRC_BIT_REG))
        'Call SetReadData2OTPCat_byOTPIdx(SearchOtpIdxByName(g_sOTP_ECID_BIT_REG))
    For lChipIdIdx = 0 To UBound(g_asChipIDName)
        Call SetReadData2OTPCat_byOTPIdx(SearchOtpIdxByName(g_asChipIDName(lChipIdIdx)))
        g_aslOTPChipReg(lChipIdIdx) = g_OTPData.Category(SearchOtpIdxByName(g_asChipIDName(lChipIdIdx))).Read.Value
    Next lChipIdIdx
    slPGMBitValue = g_OTPData.Category(SearchOtpIdxByName(g_sOTP_PRGM_BIT_REG)).Read.Value
    slECIDBitValue = g_OTPData.Category(SearchOtpIdxByName(g_sOTP_ECID_BIT_REG)).Read.Value
    slCRCBitValue = g_OTPData.Category(SearchOtpIdxByName(g_sOTP_CRC_BIT_REG)).Read.Value

    '*********************************************************************************************
    ' Datalog
    '*********************************************************************************************
    
    '== PGM BIT =='
    sLCKBitCheckTNameGrp5 = g_sOTP_PRGM_BIT_REG
    sTestName = TNameCombine(g_asLogTestName(0), g_asLogTestName(1), g_asLogTestName(2), g_asLogTestName(3), sLCKBitCheckTNameGrp5, "X", TName_OTP_X)
    If (UCase(TheExec.CurrentJob) = "QA") Or (UCase(TheExec.CurrentJob) Like "*CHAR*") Then  ' IF QA/CHAR then check lock bit must =1  2018/01/26
       Call TheExec.Flow.TestLimit(slPGMBitValue, lowVal:=1, hiVal:=1, lowCompareSign:=tlSignGreaterEqual, highCompareSign:=tlSignLessEqual, TName:=sTestName, formatStr:="%.0f")
    Else
       Call TheExec.Flow.TestLimit(slPGMBitValue, lowVal:=0, hiVal:=1, lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, TName:=sTestName, formatStr:="%.0f")
    End If
                                             
    '== ECID BIT =='
    'g_asChipIDName(0) = g_sOTP_ECID_BIT_REG
    sLCKBitCheckTNameGrp5 = g_sOTP_ECID_BIT_REG 'CStr(g_asChipIDName(0))
    sTestName = TNameCombine(g_asLogTestName(0), g_asLogTestName(1), g_asLogTestName(2), g_asLogTestName(3), sLCKBitCheckTNameGrp5, , TName_OTP_X)
    Call TheExec.Flow.TestLimit(slECIDBitValue, lowVal:=0, hiVal:=255, lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, TName:=sTestName, formatStr:="%.0f")           'TName:="Sites=" & CStr(Site) & sTestName & Rowcnt & TBD


    '== CRC BIT =='
    sLCKBitCheckTNameGrp5 = g_sOTP_CRC_BIT_REG
    sTestName = TNameCombine(g_asLogTestName(0), g_asLogTestName(1), g_asLogTestName(2), g_asLogTestName(3), sLCKBitCheckTNameGrp5, , TName_OTP_X)
    Call TheExec.Flow.TestLimit(slCRCBitValue, lowVal:=0, hiVal:=(2 ^ 8 - 1), lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, TName:=sTestName, formatStr:="%.0f")          'TName:="Sites=" & CStr(Site) & sTestName & Rowcnt & TBD

    
    '*********************************************************************************************
    ' Blank Check
    '*********************************************************************************************
    If TheExec.Flow.EnableWord("OTP_BlankCheck") = True Then '2018-12-13 Check All OTP Content =0.
        Call CheckOtpBlank(slPGMBitValue.Add(slECIDBitValue).Add(slCRCBitValue))
    End If
          
    '*********************************************************************************************
    ' Site Status Check,
    ' Define g_sbOtped, g_sbOtpedPGM, g_sbOtpedCRC, OTP_FTProg CHECK here
    '*********************************************************************************************
    Call SortOutgFlagNCheckEcidRev(slPGMBitValue, slECIDBitValue, slCRCBitValue, r_psReadPat) ', gsbOtpLocked, gsbOtpBurned)
         
    '**************************************
    ' Define the global site Variable
    ' Decide enter trim or postburn section
    '**************************************
    '___OTP_Check in CheckLCKBit '20190328
        '___In CheckLCKBit function, when all sites are burned, skip override patterns.
        '___In main flow, when all sites are burned, skip all trim section and go to postburn directly.
        '___Skip trim section when job is QA or CHAR is enabled.
    'sbSaveSiteStatus = TheExec.Sites.Selected
    'General IGXL-9.0 case
    If (UCase(TheExec.CurrentJob) = "QA") Or TheExec.EnableWord("ForceTrim_CHAR") Then
     'If g_sbOtpedPGM.All(True) Or (UCase(TheExec.CurrentJob) = "QA") Or TheExec.EnableWord("ForceTrim_CHAR") Then
        'gB_SkipTrim = True 'not used 20190720
        '*******************************
        ' FORCE SKIP TRIM FOR QA/CHAR
        '*******************************
         For Each Site In TheExec.Sites
             g_sbOtpedPGM(Site) = True
             TheExec.Sites(Site).SiteVariableValue("RunTrim") = 0
         Next Site

        
    ElseIf TheExec.EnableWord("B_Debug_ForceReTrim") Then
    'Force to disable "OTP_Enable" in the "Flow_Init_EnableWd"
        For Each Site In TheExec.Sites
                TheExec.Sites(Site).SiteVariableValue("RunTrim") = -1
        Next Site
    Else
        For Each Site In TheExec.Sites
            If g_sbOtpedPGM(Site) = False Then
                TheExec.Sites(Site).SiteVariableValue("RunTrim") = -1
            Else
                TheExec.Sites(Site).SiteVariableValue("RunTrim") = 0
            End If
        Next Site
    End If


    '___20200313, it needs the customized call inside, but we could use other way to improve here
    'Utilize OTPRev data structure to store these Default patterns information
    'Ex: g_OTPRev.Category(g_lRevIdx).Default_Pattern = "xxx_TSU_xxx"
    '___Burst default/override patterns based on the different OTP versions
    Call LoadAhbDefaultValue_basedOnOtpVersion

    
    '*******************************
    ' OTP_REVISION_ PRE-CHECK
    '*******************************
    CheckOtpRev True, "PreBurnCheck-"
    
    '*******************************
    ' CHECK REGISTERS-I2C-ADDR
    '*******************************
    '     auto_OTPCategory_GetReadDecimal g_sOTP_PRGM_BIT_REG, mSL_PGM_CHECK_BIT
    '''     auto_OTPCategory_GetReadDecimal_AHB gS_OTP_HOST_INTERFACE_I2C, RegVal(4)
    '''     '----------------------------------------------------------------------------
    '''     lActAddr = g_OTPData.Category(SearchOtpIdxByName(gS_OTP_HOST_INTERFACE_I2C)).lOtpA0
    '''     Call OTP_READREG_DSP(r_psReadPat, g_sTDI, g_sTDO, lActAddr)
    '''     Call SetReadData2OTPCat_byOTPIdx(SearchOtpIdxByName(gS_OTP_HOST_INTERFACE_I2C))
    '''
    '''     'auto_OTPCategory_SetReadDecimal gS_OTP_HOST_INTERFACE_I2C, RegVal(4) ', False
    '''     auto_OTPCategory_GetReadDecimal gS_OTP_HOST_INTERFACE_I2C, RegVal(4)
    '----------------------------------------------------------------------------
    
    '___Write ECID into AHB
    If g_bOTPFW = True Then Call GetEcidFrmOtpWr2Ahb ''''20200313, need to check if the site-loop can be removed inside

    '''    '___Datalog
    '''     sTestName = TNameCombine(g_asLogTestName(0), g_asLogTestName(1), g_asLogTestName(2), g_asLogTestName(3), "HOST-INTERFACE-SYSTEM_CONTROL-REGISTERS-I2C-ADDR", "X")
    '''     mL_Limit = g_OTPData.Category(SearchOtpIdxByName(gS_OTP_HOST_INTERFACE_I2C)).lDefaultValue
    '''
    '''     For Each Site In TheExec.Sites
    '''         If mSL_PGM_CHECK_BIT = 0 Then
    '''             Call TheExec.Flow.TestLimit(TName:=sTestName, ResultVal:=RegVal(4))
    '''         Else
    '''             Call TheExec.Flow.TestLimit(TName:=sTestName, ResultVal:=RegVal(4), lowVal:=mL_Limit, hiVal:=mL_Limit)
    '''         End If
    '''     Next Site
     
     '20190611 ___Debug only
     If (g_bDump2CsvDebugPrint) Then
        Call Dump_OTPWRData_Check_Append(sFuncName, eREGWRITE)
     End If
     
skip_default_writes:
    '___Reset Previous Active Sites
    'TheExec.Sites.Selected = sbSaveSiteStatus
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetHwCrc2Burn(Optional r_psCRCWriteTsu As PatternSet, Optional r_psCRCWrite As PatternSet, _
                            Optional r_psCRCReadTsu As PatternSet, Optional r_psCRCRead As PatternSet, _
                            Optional Validating_ As Boolean) As Long
    Dim sFuncName As String: sFuncName = "GetHwCrc2Burn"
    On Error GoTo ErrHandler
    
    If Validating_ Then
        If r_psCRCWriteTsu <> "" Then Call PrLoadPattern(r_psCRCWriteTsu.Value)
        If r_psCRCWrite <> "" Then Call PrLoadPattern(r_psCRCWrite.Value)
        If r_psCRCReadTsu <> "" Then Call PrLoadPattern(r_psCRCReadTsu.Value)
        If r_psCRCRead <> "" Then Call PrLoadPattern(r_psCRCRead.Value)
        Exit Function
    End If
    Dim asPattern() As String, lPatCnt As Long
    Dim slHWCRC As New SiteLong
    Dim slSWCRC As New SiteLong
    Dim lOtpIdx As Long
    Dim sbCheckCRC As New SiteBoolean

    If TheExec.Sites.Selected.Count = 0 Then
       TheExec.Datalog.WriteComment "*** If no Site alive, bypass " & sFuncName & " **************************************************"
       Exit Function
    End If
    
    TheExec.Datalog.WriteComment "<" + sFuncName + ">"
    
    '(2019-06-13): Do this sequence for  PTM-0,1,4 for the PTM bit
    Call CheckHwCrcEndurance_PTM
    
    '___Set CRC Value
    If TheExec.TesterMode = testModeOffline Then
        g_slOfflineCRC = GetSwCrc_ByOtpCat(eREGWRITE)
        Call auto_OTPCategory_SetWriteDecimal(g_sOTP_CRC_BIT_REG, g_slOfflineCRC)
    Else
        If r_psCRCWriteTsu.Value <> "" Then TheHdw.Patterns(r_psCRCWriteTsu).Load
        If r_psCRCWrite.Value <> "" Then TheHdw.Patterns(r_psCRCWrite).Load
        If r_psCRCReadTsu.Value <> "" Then TheHdw.Patterns(r_psCRCReadTsu).Load
        If r_psCRCRead.Value <> "" Then TheHdw.Patterns(r_psCRCRead).Load
        If r_psCRCWriteTsu.Value <> "" Then r_psCRCWriteTsu.Value = GetPatListFromPatternSet_OTP(r_psCRCWriteTsu.Value, asPattern, lPatCnt) 'OTP_CRC_MARGIN_TSU
        If r_psCRCWrite.Value <> "" Then r_psCRCWrite.Value = GetPatListFromPatternSet_OTP(r_psCRCWrite.Value, asPattern, lPatCnt) 'OTP_CRC_MARGIN_WRITE_DSC
        If r_psCRCReadTsu.Value <> "" Then r_psCRCReadTsu.Value = GetPatListFromPatternSet_OTP(r_psCRCReadTsu.Value, asPattern, lPatCnt) 'OTP_CRC_MARGIN_READ_TSU
        If r_psCRCRead.Value <> "" Then r_psCRCRead.Value = GetPatListFromPatternSet_OTP(r_psCRCRead.Value, asPattern, lPatCnt) 'OTP_CRC_MARGIN_READ_DSC
        TheHdw.PPMU.Pins(g_sDIG_PINS).Disconnect
    
        '___HW CRC
        '___Never module burst here since it will get different values every time and burn the wrong value to OTP CRC.
        
        Call CheckHwCrc(slHWCRC, 0.5)

        g_RegVal = &H0: AHB_WRITE "OTP_SLV_OTP_CMD", g_RegVal
            
            
        If g_bOTPFW = False Then  'Oneshot/multi shot need to burn 'HW CRC' into last otp address
        '___Set CRC Value
        Call auto_OTPCategory_SetWriteDecimal(g_sOTP_CRC_BIT_REG, slHWCRC)
        
        '___20200306
        '___Before writing OTP, update the write value to DSPwave first.
        lOtpIdx = SearchOtpIdxByName(g_sOTP_CRC_BIT_REG)
        For Each Site In TheExec.Sites
            Call OTPData2DSPWave(lOtpIdx)
        Next Site
        End If

        '20200406 FW : CRC has already burnt into OTP, still need to compare.
        '20200406 FW : CRC has not burnt into OTP, compare before burn CRC, not match -> bin out.
        '___SW CRC_CALC with Write Data:
        slSWCRC = GetSwCrc_ByOtpCat(eREGWRITE)
        TheExec.Flow.TestLimit slSWCRC, 0, 255, TName:="OTP_CRC_SWCRC", Unit:=unitNone
        '___Compare CRC
        sbCheckCRC = slHWCRC.Compare(EqualTo, slSWCRC)
        TheExec.Flow.TestLimit sbCheckCRC, -1, -1, TName:="OTP_CRC_CRCCheckResult", Unit:=unitNone
        

    End If
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckCrcConsistency(Optional r_psCRCWriteTsu As PatternSet, Optional r_psCRCWrite As PatternSet, _
                            Optional r_psCRCReadTsu As PatternSet, Optional r_psCRCRead As PatternSet, _
                            Optional Validating_ As Boolean) As Long

On Error GoTo ErrHandler
Dim m_sFuncName As String: m_sFuncName = "CheckCrcConsistency"
If Validating_ Then
    If r_psCRCWriteTsu <> "" Then Call PrLoadPattern(r_psCRCWriteTsu.Value)
    If r_psCRCWrite <> "" Then Call PrLoadPattern(r_psCRCWrite.Value)
    If r_psCRCReadTsu <> "" Then Call PrLoadPattern(r_psCRCReadTsu.Value)
    If r_psCRCRead <> "" Then Call PrLoadPattern(r_psCRCRead.Value)
    Exit Function
End If
Dim asPattern() As String, mL_PatCount As Long
Dim slHWCRC  As New SiteLong   'HW CRC_CALC
Dim slOTPCRC As New SiteLong   'OTP BurnCRC
Dim slSWCRC  As New SiteLong   'SW CRC_CALC
Dim sbCheckCRC As New SiteBoolean

    If g_sOTPRevisionType = "OTP_V01" Then
        TheExec.Datalog.WriteComment "<" + m_sFuncName + ">" & ":SKIP CRC CHECK for OTP_V1(Only ECID)"
    Exit Function
    End If
       
    If TheExec.Sites.Selected.Count = 0 Then
       TheExec.Datalog.WriteComment "*** If no Site alive, bypass " & m_sFuncName & " **************************************************"
       Exit Function
    End If
    
 TheHdw.StartStopwatch 'Timer start
 
    
    If r_psCRCWriteTsu.Value <> "" Then TheHdw.Patterns(r_psCRCWriteTsu).Load
    If r_psCRCWrite.Value <> "" Then TheHdw.Patterns(r_psCRCWrite).Load
    If r_psCRCReadTsu.Value <> "" Then TheHdw.Patterns(r_psCRCReadTsu).Load
    If r_psCRCRead.Value <> "" Then TheHdw.Patterns(r_psCRCRead).Load
    If r_psCRCWriteTsu.Value <> "" Then r_psCRCWriteTsu.Value = GetPatListFromPatternSet_OTP(r_psCRCWriteTsu.Value, asPattern, mL_PatCount) 'OTP_CRC_MARGIN_TSU
    If r_psCRCWrite.Value <> "" Then r_psCRCWrite.Value = GetPatListFromPatternSet_OTP(r_psCRCWrite.Value, asPattern, mL_PatCount) 'OTP_CRC_MARGIN_WRITE_DSC
    If r_psCRCReadTsu.Value <> "" Then r_psCRCReadTsu.Value = GetPatListFromPatternSet_OTP(r_psCRCReadTsu.Value, asPattern, mL_PatCount) 'OTP_CRC_MARGIN_READ_TSU
    If r_psCRCRead.Value <> "" Then r_psCRCRead.Value = GetPatListFromPatternSet_OTP(r_psCRCRead.Value, asPattern, mL_PatCount) 'OTP_CRC_MARGIN_READ_DSC
  
    TheHdw.PPMU.Pins(g_sDIG_PINS).Disconnect
    '___A.HW CRC:
        Call CheckHwCrc(slHWCRC, 0.01) 'CRC_Integrity_Check

    '___B.SW CRC_CALC with Write Data:
        If g_bOTPFW = False Or g_bFWDlogCheck = True Then
            slSWCRC = GetSwCrc_ByOtpCat(eREGREAD)
'        Else
'            slSWCRC = GetSwCrc_ByOtpCat(eREGWRITE)
'        End If
        TheExec.Flow.TestLimit slSWCRC, 0, 255, TName:="OTP_CRC_SWCRC", Unit:=unitNone
        
    '___C.OTP_Burn CRC
'        If g_bOTPFW = False Or g_bFWDlogCheck = True Then
            If TheExec.TesterMode = testModeOffline Then
                slOTPCRC = g_slOfflineCRC
            Else
                Call auto_OTPCategory_GetReadDecimal(g_sOTP_CRC_BIT_REG, slOTPCRC)
            End If
            TheExec.Flow.TestLimit slOTPCRC, 0, 255, TName:="OTP_CRC_OTPCRC", Unit:=unitNone
    
            '___Compare CRC
            sbCheckCRC = slHWCRC.Compare(EqualTo, slOTPCRC).LogicalAnd(slOTPCRC.Compare(EqualTo, slSWCRC))
'        Else
'
'            '___Compare CRC
'            sbCheckCRC = slHWCRC.Compare(EqualTo, slSWCRC)
            TheExec.Flow.TestLimit sbCheckCRC, -1, -1, TName:="OTP_CRC_CRCCheckResult", Unit:=unitNone
        End If
        
    
        
    If (g_bDump2CsvDebugPrint) Then
        Call Dump_OTPWRData_Check_Append(m_sFuncName, eREGREAD)
        End If
    
    CheckOtpRev True, "PostBurnCheck-"
    
    CheckEcid_Pst
    
Call OTP_SPT_D(" *** OTP Check CRC Consistency  , Exe Time =  ") 'Timer stop
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + m_sFuncName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'20190526 Compare OTP-to be written values and AHB_read back values
'___Checking by the conditions defined in syntax
Public Function CheckAhbNOtpWrite(Optional CompType As g_eAHB_OTP_COMP_TYPE = g_eAHB_OTP_COMP_TYPE.eCHECK_ALL)
    Dim sFuncName As String: sFuncName = "CheckAhbNOtpWrite"
    On Error GoTo ErrHandler
    
    Dim lIdx As Long, lChkCnt As Long
    
    Dim aslWriteOTP() As New SiteLong, aslReadAHBOTP() As New SiteLong, aslCalAhOtp() As New SiteLong
    Dim asOtpRegisterName() As String, asRegisterName() As String, asvChkResult() As New SiteVariant, asBWOffset() As String
    Dim lOtpRegOfs As Long, lBitWidth As Long
    Dim asOtpOwner() As String, asDefaultorReal() As String
    
    Select Case CompType
        Case g_eAHB_OTP_COMP_TYPE.eCHECK_BY_CONDITION
            lChkCnt = g_DictOTPPreCheckIndex.Count
        Case g_eAHB_OTP_COMP_TYPE.eCHECK_ALL
            lChkCnt = g_Total_OTP ''''20200313, was UBound(g_OTPData.Category) + 1, because of New AHB method
    End Select
    
    ReDim aslWriteOTP(lChkCnt - 1): ReDim aslReadAHBOTP(lChkCnt - 1): ReDim aslCalAhOtp(lChkCnt - 1)
    ReDim asOtpRegisterName(lChkCnt - 1): ReDim asRegisterName(lChkCnt - 1): ReDim asvChkResult(lChkCnt - 1): ReDim asBWOffset(lChkCnt - 1)
    ReDim asOtpOwner(lChkCnt - 1): ReDim asDefaultorReal(lChkCnt - 1) As String
    
    Dim lKey               As Long
    Dim sComment           As String
    Dim lOtpIdx As Long
    Dim sbSaveSiteStatus As New SiteBoolean
    
    Dim slTempAHBReadVal As New SiteLong
    Dim sAHBRegOtpIdx As String
    Dim asIdxArr() As String
    Dim sbCheckAhbNOtpWrite As New SiteBoolean

    If TheExec.Sites.Selected.Count = 0 Then Exit Function
    TheExec.Datalog.WriteComment "<FuncName> " + sFuncName + ": Please check these OTP Category Item.@ InstanceName=" + TheExec.DataManager.InstanceName
    '___Site status check

'    TheExec.Sites.Selected = g_sbOtpedPGM
    sbSaveSiteStatus = TheExec.Sites.Selected
    For Each Site In TheExec.Sites.Selected '.Existing
        If TheExec.Sites(Site).SiteVariableValue("RunTrim") = -1 Then
            sbCheckAhbNOtpWrite(Site) = True
        End If
    Next Site
    TheExec.Sites.Selected.Value = sbCheckAhbNOtpWrite
    If sbCheckAhbNOtpWrite.All(False) = True Then GoTo Exit_Function
    
    
        
    '___Get All AHB Read data into OTPData whenever OTPData has the corresponding AHB register name
    Call ReadAhbToCat
        
    Select Case CompType
        Case g_eAHB_OTP_COMP_TYPE.eCHECK_ALL
            TheExec.Datalog.WriteComment "** <<Check All AHB-OTP>> **"
            
            
            For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
                  With g_OTPData.Category(lOtpIdx)
                       asOtpRegisterName(lOtpIdx) = .sOtpRegisterName
                       asRegisterName(lOtpIdx) = .sRegisterName
                       asOtpOwner(lOtpIdx) = .sOTPOwner
                       asDefaultorReal(lOtpIdx) = .sDefaultORReal
                       lBitWidth = .lBitWidth
                       lOtpRegOfs = .lOtpRegOfs
                       asBWOffset(lOtpIdx) = lBitWidth
                 End With
                    '___Get OTP write(to-be-Burned) data
                    aslWriteOTP(lOtpIdx) = g_OTPData.Category(lOtpIdx).Write.Value

                    If TheExec.TesterMode = testModeOffline Then
                        If g_OTPData.Category(lOtpIdx).sAhbAddress <> "NA" Then
                            g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs = g_OTPData.Category(lOtpIdx).Write.Value
                            sAHBRegOtpIdx = g_dictAHBRegToOTPDataIdx.Item(asRegisterName(lOtpIdx)) '20190717
                            asIdxArr = Split(sAHBRegOtpIdx, ",")
                            If lOtpIdx = asIdxArr(UBound(asIdxArr)) Then
                                With g_OTPData.Category(lOtpIdx)
                                    For Each Site In TheExec.Sites
                                        slTempAHBReadVal(Site) = CLng(.Write.Value)
                                    Next Site
                                    '.svAhbReadVal(Site) = slTempAHBReadVal.ShiftRight(lOtpRegOfs)
                                    .svAhbReadVal(Site) = slTempAHBReadVal.ShiftLeft(lOtpRegOfs)
                                    If lOtpIdx <> asIdxArr(0) Then
                                        'For otpidx_of_singleAHB = 0 To UBound(asIdxArr)
                                        g_OTPData.Category(lOtpIdx).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadVal.Add(g_OTPData.Category(lOtpIdx - 1).svAhbReadVal)
                                        'update previous .svAhbReadVal
                                        g_OTPData.Category(lOtpIdx - 1).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadVal
                                        'Next otpidx_of_singleAHB
                                    Else
                                        .svAhbReadVal = .svAhbReadValByMaskOfs
                                    End If
                                    '.svAhbReadValByMaskOfs(Site) = (.Write.Value(Site)) And (.lCalDeciAhbByMaskOfs)
                                End With
                            Else
                                g_OTPData.Category(lOtpIdx).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs
                            End If
                                aslReadAHBOTP(lOtpIdx) = g_OTPData.Category(lOtpIdx).svAhbReadVal
                                aslCalAhOtp(lOtpIdx) = g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs
                                asvChkResult(lOtpIdx) = aslCalAhOtp(lOtpIdx).Compare(EqualTo, aslWriteOTP(lOtpIdx))
                        End If
                    Else 'If TheExec.TesterMode = testModeOnline Then
                    
                        If g_OTPData.Category(lOtpIdx).sAhbAddress <> "NA" Then
                            
                            aslReadAHBOTP(lOtpIdx) = g_OTPData.Category(lOtpIdx).svAhbReadVal
                            aslReadAHBOTP(lOtpIdx) = aslReadAHBOTP(lOtpIdx).ShiftRight(lOtpRegOfs)
                            '___20200313, check if it can remove Site-Loop
                            If (False) Then 'If YES, change it to True
                                aslCalAhOtp(lOtpIdx) = aslReadAHBOTP(lOtpIdx).BitWiseAnd(g_OTPData.Category(lOtpIdx).lCalDeciAhbByMaskOfs)
                            Else
                                For Each Site In TheExec.Sites
                                    aslCalAhOtp(lOtpIdx)(Site) = (aslReadAHBOTP(lOtpIdx)(Site)) And (g_OTPData.Category(lOtpIdx).lCalDeciAhbByMaskOfs)
                                Next Site
                            End If
                            asvChkResult(lOtpIdx) = aslCalAhOtp(lOtpIdx).Compare(EqualTo, aslWriteOTP(lOtpIdx))
                            
                        Else
                            aslReadAHBOTP(lOtpIdx) = 0
                            aslCalAhOtp(lOtpIdx) = 0
                            'mSB_Check = True
                            asvChkResult(lOtpIdx) = -999
                        End If
                    End If
            Next lOtpIdx
            
    Case g_eAHB_OTP_COMP_TYPE.eCHECK_BY_CONDITION

        TheExec.Datalog.WriteComment "** <<Check AHB-OTP By Condition>> **"
        lIdx = 0
         
        For lKey = 1 To g_DictOTPPreCheckIndex.Count
            lOtpIdx = g_DictOTPPreCheckIndex.Item(lKey)
                
            With g_OTPData.Category(lOtpIdx)
                asOtpRegisterName(lKey - 1) = .sOtpRegisterName
                asRegisterName(lKey - 1) = .sRegisterName
                asOtpOwner(lKey - 1) = .sOTPOwner
                asDefaultorReal(lKey - 1) = .sDefaultORReal
                lBitWidth = .lBitWidth
                lOtpRegOfs = .lOtpRegOfs
                asBWOffset(lKey - 1) = lBitWidth
            End With
            '___Get OTP write(to-be-Burned) data
            aslWriteOTP(lKey - 1) = g_OTPData.Category(lOtpIdx).Write.Value

            If TheExec.TesterMode = testModeOffline Then
                If g_OTPData.Category(lOtpIdx).sAhbAddress <> "NA" Then
                    g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs = g_OTPData.Category(lOtpIdx).Write.Value
                    sAHBRegOtpIdx = g_dictAHBRegToOTPDataIdx.Item(asRegisterName(lKey - 1)) '20190717
                    asIdxArr = Split(sAHBRegOtpIdx, ",")
                    If lOtpIdx = asIdxArr(UBound(asIdxArr)) Then
                    
                        With g_OTPData.Category(lOtpIdx)
                            '___20200313, check if it can remove Site-Loop
                            If (False) Then 'If YES, change it to True
                                slTempAHBReadVal = .Write.Value
                                .svAhbReadVal = slTempAHBReadVal.ShiftLeft(lOtpRegOfs)
                            Else
                                For Each Site In TheExec.Sites
                                    slTempAHBReadVal(Site) = CLng(.Write.Value)
                                    '.svAhbReadVal(Site) = slTempAHBReadVal.ShiftRight(lOtpRegOfs)
                                    .svAhbReadVal(Site) = slTempAHBReadVal.ShiftLeft(lOtpRegOfs)
                                Next Site
                            End If

                            If lOtpIdx <> asIdxArr(0) Then
                                'For otpidx_of_singleAHB = 0 To UBound(asIdxArr)
                                g_OTPData.Category(lOtpIdx).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadVal.Add(g_OTPData.Category(lOtpIdx - 1).svAhbReadVal)
                                'update previous .svAhbReadVal
                                g_OTPData.Category(lOtpIdx - 1).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadVal
                                'Next otpidx_of_singleAHB
                            Else
                                .svAhbReadVal = .svAhbReadValByMaskOfs
                            End If
                                '.svAhbReadValByMaskOfs(Site) = (.Write.Value(Site)) And (.lCalDeciAhbByMaskOfs)
                        End With
                    Else
                        g_OTPData.Category(lOtpIdx).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs
                    End If
                    aslReadAHBOTP(lKey - 1) = g_OTPData.Category(lOtpIdx).svAhbReadVal
                    aslCalAhOtp(lKey - 1) = g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs
                    asvChkResult(lKey - 1) = aslCalAhOtp(lKey - 1).Compare(EqualTo, aslWriteOTP(lKey - 1))
                End If
            Else ' If TheExec.TesterMode = testModeonline Then
                If g_OTPData.Category(lOtpIdx).sAhbAddress <> "NA" Then
                    
                    aslReadAHBOTP(lKey - 1) = g_OTPData.Category(lOtpIdx).svAhbReadVal
                    aslReadAHBOTP(lKey - 1) = aslReadAHBOTP(lKey - 1).ShiftRight(lOtpRegOfs)
                    '___20200313, check if it can remove Site-Loop
                    If (False) Then 'If YES, change it to True
                        aslCalAhOtp(lKey - 1) = aslReadAHBOTP(lKey - 1).BitWiseAnd(g_OTPData.Category(lOtpIdx).lCalDeciAhbByMaskOfs)
                    Else
                        For Each Site In TheExec.Sites
                            aslCalAhOtp(lKey - 1)(Site) = (aslReadAHBOTP(lKey - 1)(Site)) And (g_OTPData.Category(lOtpIdx).lCalDeciAhbByMaskOfs)
                        Next Site
                    End If

                    asvChkResult(lKey - 1) = aslCalAhOtp(lKey - 1).Compare(EqualTo, aslWriteOTP(lKey - 1))
                    
                Else
                    aslReadAHBOTP(lKey - 1) = 0
                    aslCalAhOtp(lKey - 1) = 0
                    'mSB_Check = True
                    asvChkResult(lKey - 1) = -999
                End If
            End If
        Next lKey
    End Select
    
    'B).DataLog
    Dim sTemp As String
    Dim slFailCnt As New SiteLong
                   
    TheExec.Datalog.WriteComment "<" + sFuncName + ">, InstanceName:" & TheExec.DataManager.InstanceName
    TheExec.Datalog.WriteComment "<" + sFuncName + ">, AHB CHECK CNT       :" & lChkCnt
    If CompType = g_eAHB_OTP_COMP_TYPE.eCHECK_ALL Then
       TheExec.Datalog.WriteComment "<" + sFuncName + ">, CHECK All OTP-AHB"
    Else
       TheExec.Datalog.WriteComment "<" + sFuncName + ">, CHECK OTPOwner      :" & g_sOTP_OWNER_FOR_CHECK
       TheExec.Datalog.WriteComment "<" + sFuncName + ">, CHECK DefaultORReal :" & g_sOTP_DEFAULT_REAL_FOR_PRECHECK
    End If
    TheExec.Datalog.WriteComment "<" + sFuncName + ">, PLEASE CHECK THESE OTP/AHB LISTS!"

       
    For Each Site In TheExec.Sites.Selected
        sTemp = "" 'FormatLog("Comment", -20)
        sTemp = " [Site" & CStr(Site) & "]"
        sTemp = sTemp & "," & FormatLog("OTP-REGName", -g_lOTPCateNameMaxLen - 17) & "," & FormatLog("OTP-TrimCode", -15)
        sTemp = sTemp & "," & FormatLog("AHB-REGName", -g_lOTPCateNameMaxLen) & "," & FormatLog("AHB-TrimCode", -15)
        sTemp = sTemp & "," & FormatLog("Check", -10)
        sTemp = sTemp & "," & FormatLog("BW", -5)
        sTemp = sTemp & "," & FormatLog("OTPOWNER", -10)
        sTemp = sTemp & "," & FormatLog("Default|Real", -15)
        
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "*********************************************************************************************** " & _
                    "*********************************************************************************************************"
        TheExec.Datalog.WriteComment sTemp
        TheExec.Datalog.WriteComment "*********************************************************************************************** " & _
                    "*********************************************************************************************************"
        slFailCnt(Site) = 0
        For lIdx = 0 To UBound(asOtpRegisterName)
            If InStr(UCase(g_asBYPASS_AHBOTP_CHECK), UCase(asOtpRegisterName(lIdx))) = 0 Then
                sComment = ""
            Else
                sComment = "(NoNeedToSetOTP)"
            End If
            
            sTemp = ""
            sTemp = " [Site" & CStr(Site) & "]"
            'sTemp = sTemp & "," & FormatLog(asOtpRegisterName(lIdx), -g_lOTPCateNameMaxLen) & "," & FormatLog("&H" + Hex(aslWriteOTP(lIdx)), -12)
            'sTemp = sTemp & "," & FormatLog(asRegisterName(lIdx), -g_lOTPCateNameMaxLen - 10) & "," & FormatLog("&H" + Hex(aslReadAHBOTP(lIdx)), -12)
            sTemp = sTemp & "," & FormatLog(asOtpRegisterName(lIdx), -g_lOTPCateNameMaxLen) & "," & FormatLog(sComment, -17) & "," & FormatLog(aslWriteOTP(lIdx), -15)
            sTemp = sTemp & "," & FormatLog(asRegisterName(lIdx), -g_lOTPCateNameMaxLen) & "," & FormatLog(aslCalAhOtp(lIdx), -15)
            sTemp = sTemp & "," & FormatLog(asvChkResult(lIdx), -10)
            sTemp = sTemp & "," & FormatLog(asBWOffset(lIdx), -5)
            sTemp = sTemp & "," & FormatLog(asOtpOwner(lIdx), -10)
            sTemp = sTemp & "," & FormatLog(asDefaultorReal(lIdx), -15)
            
            
            If asvChkResult(lIdx) <> -1 And (InStr(UCase(g_asBYPASS_AHBOTP_CHECK), UCase(asOtpRegisterName(lIdx))) = 0) Then slFailCnt(Site) = slFailCnt(Site) + 1
            
            'Datalog for check fail only
            If asvChkResult(lIdx) <> -1 Then
                TheExec.Datalog.WriteComment sTemp
            End If
        Next lIdx
        TheExec.Datalog.WriteComment ""
    Next Site
    
    '___DATALOG:
    Dim sTName As String
    sTName = "AHBOTPWritePreCheck-slFailCnt"

    '___20200313, copy from MP7P, Need to check here.
    If (True) Then
        TheExec.Flow.TestLimit slFailCnt, 0, 0, TName:=sTName
    Else
        ''''Original, it's masked out here
        'If g_bOtpEnable = True Then
        '    ___User Maintain high Limit
        '    TheExec.Flow.TestLimit slFailCnt, 0, 0, TName:=sTName
        '
        'Else
        '    TheExec.Flow.TestLimit slFailCnt, 0, 999, TName:=sTName
        'End If
    End If
    

Exit_Function:
TheExec.Sites.Selected = sbSaveSiteStatus

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function GetEcidFrmOtpWr2Ahb()
    Dim sFuncName As String: sFuncName = "GetEcidFrmOtpWr2Ahb"
    On Error GoTo ErrHandler
    Dim sTName As String
    Dim slPGMBitValue As New SiteLong
    'Dim OTP_CHIP_ID_Name(7) As String
    Dim slChipIdReadVal()  As New SiteLong
    Dim slChipIdWriteVal()  As New SiteLong
    Dim slChipIdSum         As New SiteLong
    Dim lChipIdIdx As Long
    Dim sbSaveSiteStatus As New SiteBoolean
    
    ReDim slChipIdReadVal(UBound(g_asChipIDName))
    ReDim slChipIdWriteVal(UBound(g_asChipIDName))

    'Call auto_OTPCategory_GetReadDecimal(g_sOTP_PRGM_BIT_REG, slPGMBitValue)    ', False)
     slChipIdSum = 0
      For lChipIdIdx = 0 To UBound(g_asChipIDName)
          Call auto_OTPCategory_GetWriteDecimal(CStr(g_asChipIDName(lChipIdIdx)), slChipIdWriteVal(lChipIdIdx)) ', False)
          Call auto_OTPCategory_GetReadDecimal(CStr(g_asChipIDName(lChipIdIdx)), slChipIdReadVal(lChipIdIdx)) ', False)
          slChipIdSum = slChipIdSum.Add(slChipIdReadVal(lChipIdIdx))
      Next lChipIdIdx
    
    sTName = "OTP_ECID_CHECK_PRE_" & "ECID-SUM"
    Call TheExec.Flow.TestLimit(TName:=sTName, ResultVal:=slChipIdSum)
    
    '20200217
    sbSaveSiteStatus = TheExec.Sites.Selected
    '(1)The ECID has been burned, Set ECID Data from OTP Read:
    TheExec.Sites.Selected = g_sbOtpedECID
    If TheExec.Sites.Selected.Count = 0 Then GoTo SaveSiteStatus
        For lChipIdIdx = 0 To UBound(g_asChipIDName)
            g_RegVal = slChipIdReadVal(lChipIdIdx)
            Call auto_OTPCategory_SetWriteDecimal_AHB(CStr(g_asChipIDName(lChipIdIdx)), g_RegVal) ', False)
            '___20200313, check if it needs, because last one (above) already SetWrite on OTP's Write structure and AHBWrite inside
            Call auto_OTPCategory_SetWriteDecimal(CStr(g_asChipIDName(lChipIdIdx)), g_RegVal) ', False) 'Need to Check this steps.(2018/06/22) Write0 or write ECID data again
            'Call auto_OTPCategory_GetReadDecimal_AHB(CStr(g_asChipIDName(lChipIdIdx)), g_RegVal) ', False)
        Next lChipIdIdx
SaveSiteStatus:
     TheExec.Sites.Selected = sbSaveSiteStatus

     '(2)Do the ECID fresh part
     'If g_sbOtpedECID.All(False) Then Exit Function
     
    TheExec.Sites.Selected = g_sbOtpedECID.LogicalNot
    If TheExec.Sites.Selected.Count = 0 Then GoTo Exit_Function
    'ECID Data from Prober:
    For lChipIdIdx = 0 To UBound(g_asChipIDName)
        g_RegVal = slChipIdWriteVal(lChipIdIdx)
        Call auto_OTPCategory_SetWriteDecimal_AHB(CStr(g_asChipIDName(lChipIdIdx)), g_RegVal) ' ,False)
        'Call auto_OTPCategory_GetReadDecimal_AHB(CStr(g_asChipIDName(lChipIdIdx)), g_RegVal) ', False)
        Call auto_OTPCategory_GetWriteDecimal(CStr(g_asChipIDName(lChipIdIdx)), g_RegVal) ', True)
    Next lChipIdIdx
Exit_Function:
    TheExec.Sites.Selected = sbSaveSiteStatus
       
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190611
Public Function CheckEcid(psReadPat As PatternSet)
    Dim sFuncName As String: sFuncName = "CheckEcid"
    On Error GoTo ErrHandler
    
    Dim lActAddr   As Long
    Dim dVDDVVal As String
    Dim asTestName() As String
    Dim slOtpChipIDVal() As New SiteLong
    ReDim asTestName(UBound(g_asChipIDName)) As String
    ReDim slOtpChipIDVal(UBound(g_asChipIDName)) As New SiteLong
    Dim lECIDLastRegAddr As Long
    Dim lIdx As Long
    
    'Datalog Naming:
    dVDDVVal = "VDD" & Format(TheHdw.DCVI.Pins(g_sVDD_PINNAME).Voltage, "0.#0") & "V"
    g_asLogTestName(0) = "OTP"
    g_asLogTestName(1) = "ECID"  ' SubTestMode
    g_asLogTestName(2) = "CHECK"
    g_asLogTestName(3) = "OTPOffset-&H" & Format(Hex(g_iOTP_ADDR_OFFSET), "000") ' "X"
    g_asLogTestName(4) = "X"

    '2).Normal  Case: wafer information correct and burn other OTP    'OTP_ADDR_Offset =gC_OTP_Offset=&HE00
 
        lIdx = SearchOtpIdxByName(g_asChipIDName(UBound(g_asChipIDName)))
        lECIDLastRegAddr = Round((g_OTPData.Category(lIdx).lOtpBitStrEnd / g_iOTP_REGDATA_BW) + 0.5)
        For lActAddr = 0 To lECIDLastRegAddr
            Call OTP_READREG_DSP(psReadPat, lActAddr) ', gL_OTP_RegADDR_Read(lActAddr).Value) ', OTP_ADDR_Offset:=OTP_ADDR_Offset)
        Next lActAddr
        For lIdx = 0 To UBound(g_asChipIDName)
            'g_asLogTestName(3) = "OTPOffset-&H" & Format(Hex(OTP_ADDR_Offset), "000")
            g_asLogTestName(4) = Replace(CStr(g_asChipIDName(lIdx)), "_", "-")
            
            asTestName(lIdx) = Join(g_asLogTestName, "_")
            Call SetReadData2OTPCat_byOTPIdx(SearchOtpIdxByName(CStr(g_asChipIDName(lIdx))))
            slOtpChipIDVal(lIdx) = g_OTPData.Category(SearchOtpIdxByName(CStr(g_asChipIDName(lIdx)))).Read.Value
            
            Call TheExec.Flow.TestLimit(TName:=asTestName(lIdx), ResultVal:=slOtpChipIDVal(lIdx), formatStr:="%.0f")
        Next lIdx
       
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckOtpRev(Optional r_bDebugPrintLog As Boolean = False, _
                                        Optional r_sPreTName As String = "PreBurnCheck-")
    Dim sFuncName As String: sFuncName = "CheckOtpRev"
    On Error GoTo ErrHandler
    
    Dim asEnableLetter(2)     As String
    Dim asString(2)     As String
    Dim sOTPType         As String
    Dim sSiteDatalog      As String
    Dim aiLen(4)           As Integer
    Dim lLimit            As Long
    Dim lRevIdx      As Long
    Dim vRevName As Variant
    Dim alLogRevVal() As Long
    'Dim m_aslOTPRevVal() As New SiteLong
    Dim aslEnableRevVal() As New SiteLong
    ReDim alLogRevVal(UBound(g_asOTPRevName)) As Long
    ReDim m_aslOTPRevVal(UBound(g_asOTPRevName)) As New SiteLong
    ReDim aslEnableRevVal(UBound(g_asOTPRevName)) As New SiteLong


    If g_sOTPRevisionType = "OTP_V01" = True Then Exit Function
    
    TheExec.Datalog.WriteComment "<" & TheExec.DataManager.InstanceName & ">"
    TheExec.Datalog.WriteComment "sFuncName:" & sFuncName

    Dim psAhbSel As New PatternSet  '20190415
    psAhbSel.Value = g_sAHB_SEL
    GetPatListNExecutePat psAhbSel
    
    
'    If TheExec.TesterMode = testModeOffline Then
'        If r_sPreTName Like "*Pre*" Then
'            For Each Site In theexec.Sites
'                For lRevIdx = 0 To UBound(g_asOTPRevName)
'                    auto_OTPCategory_GetWriteDecimal CStr(g_asOTPRevName(lRevIdx)), aslEnableRevVal(lRevIdx)  ', False
'                    g_OTPData.Category(SearchOtpIdxByName(g_asOTPRevName(lRevIdx))).Read.Value = aslEnableRevVal(lRevIdx)
'                Next lRevIdx
'            Next Site
'        Else
'            For lRevIdx = 0 To UBound(g_asOTPRevName)
'                auto_OTPCategory_GetReadDecimal_AHB CStr(g_asOTPRevName(lRevIdx)), m_aslOTPRevVal(lRevIdx) ', True
'            Next lRevIdx
'        End If
'    Else
        If r_sPreTName Like "*Pre*" Then
            
            '20190514 Add setwrite procedure to set the Revision/Consumer/Platform
            '___Update Rev info into AHB
            If g_bOTPFW = True Or g_bOTPcmpAHB = True Then
                '********************User Maintain**************************
                auto_OTPCategory_GetWriteDecimal g_sSVN_VERSION_MSB, g_RegVal
                AHB_WRITE "OTP_SLV_OTP_ATE1", g_RegVal
                auto_OTPCategory_GetWriteDecimal g_sSVN_VERSION_LSB, g_RegVal
                AHB_WRITE "OTP_SLV_OTP_ATE2", g_RegVal
        
                auto_OTPCategory_GetWriteDecimal g_sOTP_TPVERSION_S, g_RegVal
                AHB_WRITE "OTP_SLV_OTP_ATE0.MINOR_OTP_VERSION", g_RegVal
        
                auto_OTPCategory_GetWriteDecimal g_sOTP_TPVERSION_M, g_RegVal
                AHB_WRITE "OTP_SLV_OTP_ATE0.MAJOR_OTP_VERSION", g_RegVal
                '********************User Maintain**************************
            End If
            
            '___20200313, remove the Site-Loop as MP7P
            Dim lOtpIdx As Long
            
            ''For Each Site In TheExec.Sites
                'If g_sbOtpedECID(Site) = False Then
                    'Read From default(Enable OTP Version)
                    'For lRevIdx = 0 To UBound(g_asOTPRevName)
                        'aslEnableRevVal(lRevIdx) = g_OTPData.Category(SearchOtpIdxByName(g_asOTPRevName(lRevIdx))).Write.Value
                    'Next lRevIdx
                'Else
                    'Read From OTP
                    For lRevIdx = 0 To UBound(g_asOTPRevName)
                        lOtpIdx = SearchOtpIdxByName(g_asOTPRevName(lRevIdx)) '___20200313, by this way to save 2 statement search in Write/Read
                        Call SetReadData2OTPCat_byOTPIdx(lOtpIdx)
                        aslEnableRevVal(lRevIdx) = g_OTPData.Category(lOtpIdx).Write.Value
                        m_aslOTPRevVal(lRevIdx) = g_OTPData.Category(lOtpIdx).Read.Value
                    Next lRevIdx
                'End If
            ''Next Site

        Else
            'Read From AHB
            For lRevIdx = 0 To UBound(g_asOTPRevName)
                aslEnableRevVal(lRevIdx) = g_OTPData.Category(SearchOtpIdxByName(g_asOTPRevName(lRevIdx))).Write.Value
                auto_OTPCategory_GetReadDecimal_AHB CStr(g_asOTPRevName(lRevIdx)), m_aslOTPRevVal(lRevIdx) ', True
            Next lRevIdx
        End If
'    End If
 
    'If (CHECK) Then
    'If theexec.EnableWord("OTP_REVISION_CHK") = True Then
        'Define it based on g_asOTPRevName array
        vRevName = Array("OTP_REVISION_TYPE", "OTP_CONSUMER_TYPE", "PLATFORM_ID")
    
        If r_sPreTName Like "*Pre*" Then
            For Each Site In TheExec.Sites
                '20200107 Fresh die or V1 or ForceRetrim Burned die -> Don't limit bin out
                If g_sbOtpedECID(Site) = False Or (m_aslOTPRevVal(0) + m_aslOTPRevVal(1) + m_aslOTPRevVal(2)) = 0 Or TheExec.EnableWord("B_Debug_ForceReTrim") = True Then
                    TheExec.Datalog.WriteComment "OTP_REVERSION_CHECK : Site" & CStr(Site) & " is Fresh die w/o ECID"
                    For lRevIdx = 0 To UBound(g_asOTPRevName)
                        m_TestName = "OTP_ECID_CHECK_REVISION_" & Replace(r_sPreTName & vRevName(lRevIdx), "_", "-")
                        TheExec.Flow.TestLimit aslEnableRevVal(lRevIdx), TName:=m_TestName, formatStr:="%.0f"
                    Next lRevIdx
                Else
                    TheExec.Datalog.WriteComment "OTP_REVERSION_CHECK : Site" & CStr(Site) & " ECID from OTP"
                    For lRevIdx = 0 To UBound(g_asOTPRevName)
                        lLimit = m_aslOTPRevVal(lRevIdx)
                        m_TestName = "OTP_ECID_CHECK_REVISION_" & Replace(r_sPreTName & vRevName(lRevIdx), "_", "-")
                        If m_aslOTPRevVal(0) + m_aslOTPRevVal(1) + m_aslOTPRevVal(2) = 0 Then 'If =0, means the 1st time choose V1 to burn ECID only
                            TheExec.Flow.TestLimit aslEnableRevVal(lRevIdx), TName:=m_TestName, formatStr:="%.0f"
                        Else
                            TheExec.Flow.TestLimit aslEnableRevVal(lRevIdx), TName:=m_TestName, formatStr:="%.0f", lowVal:=lLimit, hiVal:=lLimit
                        End If
                    Next lRevIdx
                End If
            Next Site
        Else
            'PostBurnCheck @EnableWord("OTP_REVISION_CHECK") = True '2018/08/14
            TheExec.Datalog.WriteComment " EnableWord:OTP_REVISION_CHECK = " & Replace(TheExec.EnableWord("OTP_REVISION_CHECK"), "-1", "True")
            For lRevIdx = 0 To UBound(g_asOTPRevName)
                lLimit = g_OTPData.Category(SearchOtpIdxByName(g_asOTPRevName(lRevIdx))).lDefaultValue
                m_TestName = "OTP_ECID_CHECK_REVISION_" & Replace(r_sPreTName & vRevName(lRevIdx), "_", "-")
                TheExec.Flow.TestLimit m_aslOTPRevVal(lRevIdx), TName:=m_TestName, formatStr:="%.0f", lowVal:=lLimit, hiVal:=lLimit
            Next lRevIdx
        End If
        
    'End If
    
    If (r_bDebugPrintLog) Then
    
        sOTPType = g_sOTPType
        
        TheExec.Datalog.WriteComment ("")
        TheExec.Datalog.WriteComment (String(100, "*"))
        TheExec.Datalog.WriteComment ("*" & Space(98) & "*")
        TheExec.Datalog.WriteComment "*" & Space(4) & "FuncName : " & sFuncName & FormatLog("*", 100 - Len("*" & Space(4) & "FuncName : " & sFuncName))
        TheExec.Datalog.WriteComment ("*" & Space(98) & "*")
        For Each Site In TheExec.Sites
            sOTPType = ""
            
            asEnableLetter(2) = ConvertOtpVersion_Value2String(ePLATFORM_ID, aslEnableRevVal(2)(Site), asString(2))
            asEnableLetter(1) = ConvertOtpVersion_Value2String(eOTP_CONSUMER_TYPE, aslEnableRevVal(1)(Site), asString(1))
            asEnableLetter(0) = ConvertOtpVersion_Value2String(eOTP_REVISION_TYPE, aslEnableRevVal(0)(Site), asString(0))
        
           'PLATFORM_ID + OTP_CONSUMER_TYPE +OTP_REVISION_TYPE
            sOTPType = asEnableLetter(2) & asEnableLetter(1) & asEnableLetter(0)
            sSiteDatalog = r_sPreTName & "Site(" & Site & ") "
            'If g_sbOtped(Site) = False Then
                'For lRevIdx = 0 To UBound(g_asOTPRevName)
                    'alLogRevVal(lRevIdx) = aslEnableRevVal(lRevIdx)
                'Next lRevIdx
            'Else
                'For lRevIdx = 0 To UBound(g_asOTPRevName)
                    'alLogRevVal(lRevIdx) = m_aslOTPRevVal(lRevIdx)
                'Next lRevIdx
            'End If
            
            aiLen(4) = 100 - Len("*" & Space(60) & "Enabled Version" & Space(8) & "OTPed Version")
            aiLen(3) = 100 - Len("*    " & sSiteDatalog & "OTP_Type              : " & sOTPType)
            aiLen(2) = 100 - Len("*    " & sSiteDatalog & "PLATFORM_ID           : " & asEnableLetter(2) & " <<= " & FormatLog(asString(2), -12) & " =d'" & CStr(aslEnableRevVal(2)(Site)) & Space(15) & "  d'" & CStr(m_aslOTPRevVal(2)(Site)))
            aiLen(1) = 100 - Len("*    " & sSiteDatalog & "OTP_CONSUMER_TYPE     : " & asEnableLetter(1) & " <<= " & FormatLog(asString(1), -12) & " =d'" & CStr(aslEnableRevVal(1)(Site)) & Space(15) & "  d'" & CStr(m_aslOTPRevVal(1)(Site)))
            aiLen(0) = 100 - Len("*    " & sSiteDatalog & "OTP_REVISION_TYPE     : " & asEnableLetter(0) & " <<= " & FormatLog(asString(0), -12) & " =d'" & CStr(aslEnableRevVal(0)(Site)) & Space(15) & "  d'" & CStr(m_aslOTPRevVal(0)(Site)))

            TheExec.Datalog.WriteComment ("*" & Space(60) & "Enabled Version" & Space(8) & "OTPed Version") & FormatLog("*", aiLen(4))
            TheExec.Datalog.WriteComment ("*    " & sSiteDatalog & "OTP_Type              : " & sOTPType) & FormatLog("*", aiLen(3))
            TheExec.Datalog.WriteComment ("*    " & sSiteDatalog & "PLATFORM_ID           : " & asEnableLetter(2) & " <<= " & FormatLog(asString(2), -12) & " =d'" & CStr(aslEnableRevVal(2)(Site)) & Space(15) & "  d'" & CStr(m_aslOTPRevVal(2)(Site))) & FormatLog("*", aiLen(2))
            TheExec.Datalog.WriteComment ("*    " & sSiteDatalog & "OTP_CONSUMER_TYPE     : " & asEnableLetter(1) & " <<= " & FormatLog(asString(1), -12) & " =d'" & CStr(aslEnableRevVal(1)(Site)) & Space(15) & "  d'" & CStr(m_aslOTPRevVal(1)(Site))) & FormatLog("*", aiLen(1))
            TheExec.Datalog.WriteComment ("*    " & sSiteDatalog & "OTP_REVISION_TYPE     : " & asEnableLetter(0) & " <<= " & FormatLog(asString(0), -12) & " =d'" & CStr(aslEnableRevVal(0)(Site)) & Space(15) & "  d'" & CStr(m_aslOTPRevVal(0)(Site))) & FormatLog("*", aiLen(0))

            TheExec.Datalog.WriteComment ("*" & Space(98) & "*")
        Next Site
        TheExec.Datalog.WriteComment (String(100, "*"))
    End If

    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Public Function Auto_GetECIDFromOTP(psReadPat As PatternSet) As Long
'On Error GoTo ErrHandler
'Dim sFuncName As String: sFuncName = "Auto_GetECIDFromOTP"
'
'    Dim Data As DSPWave
'    Dim mL_ADDRcnt   As Long
'    Dim ECID_read() As New SiteLong
'    ReDim ECID_read(UBound(g_asChipIDName))
'    Dim asPatArray() As String, lPatCnt As Long
'    Dim OTP_ADDR_Offset As Long
'    Dim dVDDVVal As Double
'    Dim ChipID_idx As Long
'    'Pattern3.Value = GetPatListFromPatternSet_OTP(Pattern3.Value, asPatArray, lPatCnt)
'    dVDDVVal = TheHdw.DCVI.Pins(g_sVPP_PINNAME).Voltage
'
'    TheExec.Datalog.WriteComment vbCrLf & "<funcName> " + sFuncName + ": "
'
'    For mL_ADDRcnt = 0 To 4
'        Call OTP_READREG_DSP(psReadPat, g_sTDI, g_sTDO, mL_ADDRcnt) ', gL_OTP_RegADDR_Read(mL_ADDRcnt).Value)
'    Next mL_ADDRcnt
'
'    For ChipID_idx = 0 To UBound(g_asChipIDName)
'        Call SetReadData2OTPCat_byOTPIdx(SearchOtpIdxByName(g_sOTP_ECID_BIT_REG))
'        g_aslOTPChipReg(ChipID_idx) = g_OTPData.Category(SearchOtpIdxByName(g_sOTP_ECID_BIT_REG)).Read.Value
'        'Call auto_OTPCategory_SetReadDecimal(CStr(g_asChipIDName(ChipID_idx)), g_aslOTPChipReg(ChipID_idx))
'    Next ChipID_idx
'
'    Call DecodeDeidFromOtp '20190717 Janet+Toppy+Ching
'
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function
Public Function CheckEcid_Pst()
    Dim sFuncName As String: sFuncName = "CheckEcid_Pst"
    On Error GoTo ErrHandler
    Dim sTName As String
    Dim sTestName As StandardStreamTypes
    Dim slPGMBitValue As New SiteLong
    Dim slChipIdReadVal()  As New SiteLong
    Dim slChipIdSum     As New SiteLong
    Dim asPatArray() As String, lPatCnt As Long
    Dim lChipIdIdx As Long
    Dim psReadPat  As New PatternSet
    ReDim slChipIdReadVal(UBound(g_asChipIDName)) As New SiteLong

    If g_sOTPRevisionType = "OTP_V01" Then
        TheExec.Datalog.WriteComment "<" + sFuncName + ">" & ":SKIP CHECK for OTP_V1(Only ECID)"
    Exit Function
    End If
    
    'theexec.Datalog.WriteComment "<" + sFuncName + ">"
     
    If g_bOTPFW = False Then
        '___A).Check OTP PROGRAMMED_CHECK_BIT:
        Call auto_OTPCategory_GetReadDecimal(g_sOTP_PRGM_BIT_REG, slPGMBitValue) ', True)
      
        '___B).Check OTP ECID:
        slChipIdSum = 0
        For lChipIdIdx = 0 To UBound(g_asChipIDName)
            Call auto_OTPCategory_GetReadDecimal(CStr(g_asChipIDName(lChipIdIdx)), slChipIdReadVal(lChipIdIdx))
            slChipIdSum = slChipIdSum.Add(slChipIdReadVal(lChipIdIdx))
        Next lChipIdIdx

        '___C).Check OTP I2C: 'No I2C interface in Sera
        'OTP_I2C_Name = "OTP_HOST_INTERFACE_I2C_ADDR_11"
        '
        'auto_OTPCategory_GetReadDecimal OTP_I2C_Name, g_RegVal
        'mL_Limit = g_OTPData.Category(SearchOtpIdxByName(OTP_I2C_Name)).lDefaultValue
        'sTestName = "OTP_ECID_CHECK_POST_" & Replace(OTP_I2C_Name, "_", "-")
        'TheExec.Flow.TestLimit g_RegVal, mL_Limit, mL_Limit, Unit:=unitNone, TName:=TestName
        
        '___PostCheck ProgramBit in OTP-FW mode
        '___2 ways to set read value of ProgramBit, one is reading back with DSSC-Read,
        '   the other one is force set it according to the value defined in OTP register map (by OTP version)
    Else
        '___20200313, need to understand its purpose <Notice> '20200407 JY has Checked and modified
''        Dim lOtpIdx As Long
''        lOtpIdx = SearchOtpIdxByName(g_sOTP_PRGM_BIT_REG)
''        'g_OTPData.Category(lOtpIdx).Read.Value = g_OTPRev.Category(g_lRevIdx).DefaultValue(lOtpIdx)
''        slPGMBitValue = g_OTPRev.Category(g_lRevIdx).DefaultValue(lOtpIdx)
        '___A).Check OTP PROGRAMMED_CHECK_BIT:
        psReadPat.Value = "JTG_EFUSE_READ_SEQ_DSC"
        psReadPat.Value = GetPatListFromPatternSet_OTP(psReadPat.Value, asPatArray, lPatCnt)
        Call OTP_READREG_DSP(psReadPat.Value, 0) '20200407 Read back PGM Bit From OTP
        '___Force device status as non-otped under offline mode
        For Each Site In TheExec.Sites
           slPGMBitValue = gD_wReadData.ElementLite(0)
        Next Site

        'below is the original codes
'        For Each Site In TheExec.Sites.Selected
'            g_OTPData.Category(SearchOtpIdxByName(g_sOTP_PRGM_BIT_REG)).Read.Value(Site) = g_OTPRev.Category(g_lRevIdx).DefaultValue(SearchOtpIdxByName(g_sOTP_PRGM_BIT_REG))
'        Next Site
    '___B).Check OTP ECID:
        slChipIdSum = 0
        For lChipIdIdx = 0 To UBound(g_asChipIDName)
            Call auto_OTPCategory_GetReadDecimal_AHB(CStr(g_asChipIDName(lChipIdIdx)), slChipIdReadVal(lChipIdIdx))
            slChipIdSum = slChipIdSum.Add(slChipIdReadVal(lChipIdIdx))
        Next lChipIdIdx

    End If

    sTName = "OTP_ECID_CHECK_POST_" & Replace(g_sOTP_PRGM_BIT_REG, "_", "-")
    TheExec.Flow.TestLimit slPGMBitValue, 1, 1, Unit:=unitNone, TName:=sTName
    sTName = "OTP_ECID_CHECK_POST_" & "ECID-SUM"
    TheExec.Flow.TestLimit slChipIdSum, 1, 99999, Unit:=unitNone, TName:=sTName
        
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SortOutgFlagNCheckEcidRev(r_slPGMBitValue As SiteLong, _
                                            r_slECIDBitValue As SiteLong, _
                                            r_slCRCBitValue As SiteLong, _
                                            ByVal psReadPat As PatternSet) ', _

Dim sFuncName As String: sFuncName = "SortOutgFlagNCheckEcidRev"
On Error GoTo ErrHandler
Dim slSumResult                      As New SiteLong
'Dim Result_CRC                  As New SiteLong
Dim IIdx As Integer
Dim sbSaveSiteStatus As New SiteBoolean
 '___A).OTP_doneREG_All 'g_sbOtped
        slSumResult = r_slPGMBitValue.Add(r_slECIDBitValue).Add(r_slCRCBitValue)

        For Each Site In TheExec.Sites
           If (slSumResult = 0) Then
             g_sbOtped = False
             TheExec.Datalog.WriteComment (" SITE " & Site & " : OTP IS NOT BURNED. >>> PART IS BLANK!!!!!!!")
           Else
             g_sbOtped = True
             TheExec.Datalog.WriteComment (" SITE " & Site & " : PART IS ALREADY OTPED For ECID !!!!!!!")
           End If
        Next Site
  
        m_TestName = TNameCombine(g_asLogTestName(0), g_asLogTestName(1), g_asLogTestName(2), g_asLogTestName(3), "ISPARTALREADYOTPED", , TName_OTP_X)
        For Each Site In TheExec.Sites
           If g_sbOtped = False Then
               'PART IS BLANK!
               Call TheExec.Flow.TestLimit(ResultVal:=0, lowVal:=0, hiVal:=1, lowCompareSign:=tlSignGreaterEqual, highCompareSign:=tlSignLessEqual, formatStr:="%.0f", TName:=m_TestName)
           Else
               'PART IS ALREADY OTPED.
               Call TheExec.Flow.TestLimit(ResultVal:=1, lowVal:=0, hiVal:=1, lowCompareSign:=tlSignGreaterEqual, highCompareSign:=tlSignLessEqual, formatStr:="%.0f", TName:=m_TestName)
           End If
        Next Site
        
 '___B).g_sbOtpedECID
         For Each Site In TheExec.Sites
           If (r_slECIDBitValue = 0) Then
                g_sbOtpedECID = False
                TheExec.Datalog.WriteComment (" SITE " & Site & " : OTP  IS NOT BURNED ECID!!!!!!!")
            Else
                g_sbOtpedECID = True
                TheExec.Datalog.WriteComment (" SITE " & Site & " : PART IS ALREADY OTPED ECID  !!!!!!!")
            End If
          Next Site
          
 '___C).g_sbOtpedPGM
         For Each Site In TheExec.Sites
           If (r_slPGMBitValue = 0) Then
                g_sbOtpedPGM = False
                TheExec.Datalog.WriteComment (" SITE " & Site & " : OTP  IS NOT BURNED for LockBit!!!!!!!")
            Else
                g_sbOtpedPGM = True
                TheExec.Datalog.WriteComment (" SITE " & Site & " : PART IS ALREADY OTPED for LockBit  !!!!!!!")
            End If
          Next Site
            
 '___D).g_sbOtpedCRC
         For Each Site In TheExec.Sites
            If (r_slCRCBitValue = 0) Then
                g_sbOtpedCRC = False
                'gsbOtpburned = False
                TheExec.Datalog.WriteComment (" SITE " & Site & " : OTP  IS NOT BURNED for CRC. >>> PART IS BLANK!!!!!!!")
            Else
                g_sbOtpedCRC = True
                'gsbOtpburned = True
                TheExec.Datalog.WriteComment (" SITE " & Site & " : PART IS ALREADY OTPED for CRC  !!!!!!!")
            End If
        Next Site

        Call DecodeDeidFromOtp

        
        '___G).OTP_ECID_Check:
        
        sbSaveSiteStatus = TheExec.Sites.Selected
        TheExec.Sites.Selected = g_sbOtpedECID
        If g_sbOtpedECID.All(False) = True Then
           TheExec.Sites.Selected = sbSaveSiteStatus
           Exit Function
        Else
            
             Call CheckEcid(psReadPat) '*** Check Addr.0 & Correct ECID Addr. &HE00
         
             '___F).OTP_REV_Check:
            'If TheExec.EnableWord("OTP_REVISION_CHECK") = True Then Call OTP_REV_Check(psReadPat) '*** Check Addr.0 & Correct ECID Addr. &HE00
             Call CheckOtpRev
            Dim idx As Integer
             'If ECID is burned on FTProg part, then setWrite the burned ECID.
             For idx = 0 To UBound(g_asChipIDName)
                 Call auto_OTPCategory_SetWriteDecimal(g_asChipIDName(idx), g_aslOTPChipReg(idx))
             Next idx
        End If
            
         TheExec.Sites.Selected = sbSaveSiteStatus
         
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''___OTP-AHB Compare before OTP
'Public Function CheckAhbNOtpWrite(Optional ByVal CheckAllAHB As Boolean = False)
'
'On Error GoTo ErrHandler
'Dim sFuncName As String: sFuncName = "CheckAhbNOtpWrite"
'
'Dim Addr As Long, OTPData_Idx As Long, OTPData_IdxCNT As Long
'Dim i As Long, j As Long, k As Long
'Dim mI_CategoryIndex As Long, mL_StartCategoryIndex As Long, mL_EndCategoryIndex As Long
'Dim aslWriteOTP() As New SiteLong, aslReadAHBOTP() As New SiteLong, aslCalAhOtp() As New SiteLong
'Dim mSB_Check As New SiteBoolean
'Dim asOtpRegisterName() As String, asRegisterName() As String, asvChkResult() As New SiteVariant, asBWOffset() As String
'Dim mS_CategoryName As String, mS_AHBREGName As String, lOtpRegOfs As Long, lBitWidth As Long, mL_AHBIndex As Long
'Dim mS_OTPOWNER As String, asOtpOwner() As String
'
'
'mL_StartCategoryIndex = 0
'mL_EndCategoryIndex = UBound(g_OTPData.Category)
'OTPData_IdxCNT = mL_EndCategoryIndex - mL_StartCategoryIndex + 1
'ReDim aslWriteOTP(OTPData_IdxCNT - 1): ReDim aslReadAHBOTP(OTPData_IdxCNT - 1): ReDim aslCalAhOtp(OTPData_IdxCNT - 1)
'ReDim asOtpRegisterName(OTPData_IdxCNT - 1): ReDim asRegisterName(OTPData_IdxCNT - 1): ReDim asvChkResult(OTPData_IdxCNT - 1): ReDim asBWOffset(OTPData_IdxCNT - 1)
'ReDim asOtpOwner(OTPData_IdxCNT - 1)
'
'
'If TheExec.TesterMode = testModeOffline Then
'   TheExec.Datalog.WriteComment "** Offline Mode **"
'End If
'
'If TheExec.Sites.Selected.Count = 0 Then Exit Function
'
'    '___Get All AHB Read data into OTPData whenever OTPData has the corresponding AHB register name
'    Call Auto_AHB_READDSC_New
'
'    For OTPData_Idx = 0 To UBound(g_OTPData.Category)
'          With g_OTPData.Category(OTPData_Idx)
'               mS_CategoryName = .sOtpRegisterName
'               mS_AHBREGName = .sRegisterName
'               lBitWidth = .lBitWidth
'               lOtpRegOfs = .lOtpRegOfs
'               mS_OTPOWNER = .sOtpOwner
'          End With
'
'            asOtpRegisterName(OTPData_Idx) = mS_CategoryName
'            asRegisterName(OTPData_Idx) = mS_AHBREGName
'            asOtpOwner(OTPData_Idx) = mS_OTPOWNER
'
'            '___Get OTP write(to-be-Burned) data
'            aslWriteOTP(OTPData_Idx) = g_OTPData.Category(OTPData_Idx).Write.Value
'
'            '___Get AHB Read data
'            Dim slTempAHBReadVal As New SiteLong
'            If TheExec.TesterMode = testModeOffline Then
'                '-----------------------------------------------------------------------
'                '___Offline calc only
'                Dim deci_AHB_Mask As Long
'
'                deci_AHB_Mask = 0
'                lOtpRegOfs = g_OTPData.Category(OTPData_Idx).lOtpRegOfs
'                For k = 0 To g_iAHB_BW - 1
'                deci_AHB_Mask = (g_OTPData.Category(OTPData_Idx).sAhbMask(k)) * 2 ^ (k) + deci_AHB_Mask
'                Next k
'                deci_AHB_Mask = (2 ^ (g_iAHB_BW) - 1) - deci_AHB_Mask
'                '-----------------------------------------------------------------------
'                For Each Site In TheExec.Sites
'                    With g_OTPData.Category(OTPData_Idx)
'                        slTempAHBReadVal(Site) = CLng(.Write.Value)
'                        .svAhbReadVal(Site) = slTempAHBReadVal.ShiftLeft(lOtpRegOfs)
'                        .svAhbReadValByMaskOfs(Site) = (.Write.Value(Site)) And (.lCalDeciAhbByMaskOfs)
'
'                    End With
'                Next Site
'                aslReadAHBOTP(OTPData_Idx) = g_OTPData.Category(OTPData_Idx).svAhbReadVal
'                aslCalAhOtp(OTPData_Idx) = g_OTPData.Category(OTPData_Idx).svAhbReadValByMaskOfs
'                asvChkResult(OTPData_Idx) = aslCalAhOtp(OTPData_Idx).Compare(EqualTo, aslWriteOTP(OTPData_Idx))
'            Else
'
'                If g_OTPData.Category(OTPData_Idx).sAhbAddress <> "NA" Then
'
'                    aslReadAHBOTP(OTPData_Idx) = g_OTPData.Category(OTPData_Idx).svAhbReadVal
'                    aslReadAHBOTP(OTPData_Idx) = aslReadAHBOTP(OTPData_Idx).ShiftRight(lOtpRegOfs)
'                    For Each Site In TheExec.Sites
'                        aslCalAhOtp(OTPData_Idx)(Site) = (aslReadAHBOTP(OTPData_Idx)(Site)) And (g_OTPData.Category(OTPData_Idx).lCalDeciAhbByMaskOfs)
'                    Next Site
'                    asvChkResult(OTPData_Idx) = aslCalAhOtp(OTPData_Idx).Compare(EqualTo, aslWriteOTP(OTPData_Idx))
'
'                Else
'                    aslReadAHBOTP(OTPData_Idx) = 0
'                    aslCalAhOtp(OTPData_Idx) = 0
'                    'mSB_Check = True
'                    asvChkResult(OTPData_Idx) = -999
'                End If
'            End If
'
'    Next OTPData_Idx
'
'
'     '___DataLog
'        Dim TempS As String
'        Dim FailCNT As New SiteLong
'
'            TheExec.Datalog.WriteComment "<" + sFuncName + ">, InstanceName:" & TheExec.DataManager.InstanceName
'            TheExec.Datalog.WriteComment "<" + sFuncName + ">, AHB CHECK CONDITION:" & g_sOTP_OWNER_FOR_CHECK_AHB
'            TheExec.Datalog.WriteComment "<" + sFuncName + ">, PLEASE CHECK THESE OTP/AHB LISTS!"
'
'            TempS = " [Site" & CStr(Site) & "]"
'            TempS = TempS & "," & FormatLog("OTPName", -g_lOTPCateNameMaxLen) & "," & FormatLog("GetWriteOTP", -12)
'            TempS = TempS & "," & FormatLog("REGName", -g_lOTPCateNameMaxLen - 10) & "," & FormatLog("ReadAHB", -12)
'            TempS = TempS & "," & FormatLog("[ : ]", -10)
'            TempS = TempS & "," & FormatLog("CalAHB", -10)
'            TempS = TempS & "," & FormatLog("OTPOWNER", -10) & "," & FormatLog("Check", -10)
'            TheExec.Datalog.WriteComment TempS
'
'        For Each Site In TheExec.Sites.Selected
'                FailCNT(Site) = 0
'                For OTPData_Idx = 0 To UBound(g_OTPData.Category)
'                     TempS = " [Site" & CStr(Site) & "]"
'                     TempS = TempS & "," & FormatLog(asOtpRegisterName(OTPData_Idx), -g_lOTPCateNameMaxLen) & "," & FormatLog("&H" + Hex(aslWriteOTP(OTPData_Idx)), -12)
'                     TempS = TempS & "," & FormatLog(asRegisterName(OTPData_Idx), -g_lOTPCateNameMaxLen - 10) & "," & FormatLog("&H" + Hex(aslReadAHBOTP(OTPData_Idx)), -12)
'                     TempS = TempS & "," & FormatLog(asBWOffset(OTPData_Idx), -10)
'                     TempS = TempS & "," & FormatLog("&H" + Hex(aslCalAhOtp(OTPData_Idx)), -10)
'                     TempS = TempS & "," & FormatLog(asOtpOwner(OTPData_Idx), -10) & "," & FormatLog(asvChkResult(OTPData_Idx), -10)
'
'
'                    If asvChkResult(OTPData_Idx) = False Then FailCNT(Site) = FailCNT(Site) + 1
'
'                    'g_bAHBWriteCheckDebugPrint = True 'Do not specify the flag in any function 20190527
'                    If g_bAHBWriteCheckDebugPrint = True Then
'                        TheExec.Datalog.WriteComment TempS
'                    Else
'                    '___Datalog for check fail only
'                        If asvChkResult(OTPData_Idx) = False Then
'                            If CheckAllAHB = True Then
'                                TheExec.Datalog.WriteComment TempS
'                            Else
'                                '___Datalog whenever the OTP_Owner is "System","Design",and "Trim"
'                                If InStr(UCase(g_sOTP_OWNER_FOR_CHECK_AHB), g_OTPData.Category(OTPData_Idx).sOtpOwner) Then
'                                    TheExec.Datalog.WriteComment TempS
'                                End If
'                            End If
'                        End If
'                    End If
'                Next OTPData_Idx
'
'        Next Site
'        TheExec.Datalog.WriteComment "<End of AHB-OTP-Write-Check>"
'  'auto_OTPCheckDefaulReal
'Exit_Function:
''Reset Previous Active Sites
''TheExec.Sites.Selected = sbSaveSiteStatus
'
'
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function

Private Function test_AHB_OTP_Debug() As Long 'Need to verify the purpose
    Dim sFuncName As String: sFuncName = "test_AHB_OTP_Debug"
    On Error GoTo ErrHandler
'Dim TestMode             As String
'Dim SubTestCondition     As String
'Dim TestName             As String
'Dim mL_lowVal              As Long
'Dim mL_hiVal               As Long
'
'Dim AHBScript            As String
'Dim sTPPath              As String
'
'TheExec.Datalog.WriteComment "<" + sFuncName + ">"
'
''A-AHB_SEL pattern
'         'Acore.Utilities.ACORE_ReadAnalogPatternList_And_Execute_Pattern "AHB_SEL"
'         Dim AHB_SEL As New PatternSet  '20190415
'
'    AHB_SEL.Value = g_sAHB_SEL
'    GetPatListNExecutePat AHB_SEL
'
''B-AHB_READ_WRITE_DEBUG
''Tests JTAG to AHB interface, Writes OTP_SLV_OTP_ATE0 to 0x30 and read from it
'
'
'         TheExec.Datalog.WriteComment (">>> AHB_READ_WRITE_DEBUG:")
'         '********************User Maintain**************************
''         g_RegVal = &H30
''         mL_lowVal = &H30: mL_hiVal = &H30
''         AHB_WRITE OTP_SLV_OTP_ATE0.Addr, g_RegVal
''         AHB_READ OTP_SLV_OTP_ATE0.Addr, g_RegVal
'         '********************User Maintain**************************
'
'         TestMode = "AHBOTP-Debug"
'         SubTestCondition = "_OTP-SLV-OTP-ATE0"
'         Call UniFormTestName.CreateIDSTestName: TestName = TestMode & SubTestCondition 'UniFormTestName.Customized(Group1:=TestMode, Group2:=SubTestCondition, Group3:="X", Group4:="X")
'         Call TheExec.Flow.TestLimit(TName:=TestName, ResultVal:=g_RegVal, lowVal:=mL_lowVal, hiVal:=mL_hiVal)
'
'
''C-AHB_READ_WRITE_BF_DEBUG
''Tests JTAG to AHB interface with write mask mechanism, Writes TST_CTRL_DTBO_CTRL to 0xff with mask register FABRIC_AHB_FABRIC_RMWU_WRMASK set to 0x55 and read 0x55 from TST_CTRL_DTBO_CTRL
'        TheExec.Datalog.WriteComment (">>> AHB_READ_WRITE_BF_DEBUG:")
'         g_RegVal = &HFF: AHB_WRITE TST_CTRL_DTBO_CTRL.Addr, g_RegVal
'         g_RegVal = &H55
'         AHB_WRITE FABRIC_AHB_FABRIC_RMWU_WRMASK.Addr, g_RegVal, FABRIC_AHB_FABRIC_RMWU_WRMASK.FABRIC_AHB_FABRIC_RMWU_WRMASK_RMWU_WRMASK
'         mL_lowVal = &H55: mL_hiVal = &H55
'         AHB_READDSC FABRIC_AHB_FABRIC_RMWU_WRMASK.Addr, g_RegVal
'
'          TestMode = "AHBOTP-Debug"
'          SubTestCondition = "_TST-CTRL-DTBO-CTRL"
'          Call UniFormTestName.CreateIDSTestName: TestName = TestMode & SubTestCondition ' UniFormTestName.Customized(Group1:=TestMode, Group2:=SubTestCondition, Group3:="X", Group4:="X")
'          Call TheExec.Flow.TestLimit(TName:=TestName, ResultVal:=g_RegVal, lowVal:=mL_lowVal, hiVal:=mL_hiVal)
'
'
''D.Test OTP write/Read Mechanism
''- Run sample Debug pattern EFUSE DEBUG WRITE READ  and use generic pattern  EFUSE WRITE and EFUSE_READ to write/read any OTP locations.
'            'OTP Program:
'            TheExec.Datalog.WriteComment (">>> Test OTP write/Read Mechanism:")
'            TheExec.Datalog.WriteComment ("********************** RAISING VPP to 7.5 V********************************************** ")
'            With TheHdw.DCVI(g_sVPP_PINNAME)
'                .SetVoltageAndRange 1.5 * v, 10 * v
'                 TheHdw.Wait 1 * ms
'                 .SetVoltageAndRange 4 * v, 10 * v
'                 TheHdw.Wait 1 * ms
'                  .SetVoltageAndRange 5 * v, 10 * v
'                 TheHdw.Wait 1 * ms
'                 .SetVoltageAndRange 7.5 * v, 10 * v
'             End With
'            TheExec.Datalog.WriteComment ("Force " & g_sVPP_PINNAME & " = " & Format(TheHdw.DCVI.Pins(g_sVPP_PINNAME).Voltage, "0.00"))
'
'            'sTPPath = Application.ActiveWorkbook.Path
'            AHBScript = g_sTPPath + "\DFT\TestPlan\OTP\otp_single_prog.txt"
'            Call AHBSetup(AHBScript, NEW_PATTERN_TSU1)
'
'            If False Then
'            g_RegVal = &H82: AHB_WRITE "OTP_SLV_OTP_ADDR_LO", g_RegVal
'            g_RegVal = &HF:  AHB_WRITE "OTP_SLV_OTP_ADDR_HI", g_RegVal
'
'            g_RegVal = &H5C: AHB_WRITE "OTP_SLV_OTP_DATA0", g_RegVal
'            g_RegVal = &H7F: AHB_WRITE "OTP_SLV_OTP_DATA1", g_RegVal
'            g_RegVal = &H9F: AHB_WRITE "OTP_SLV_OTP_DATA2", g_RegVal
'            g_RegVal = &H1F: AHB_WRITE "OTP_SLV_OTP_DATA3", g_RegVal
'
'            g_RegVal = &H2: AHB_WRITE "OTP_SLV_OTP_CMD", g_RegVal
'            'ANALOG_OPERATION, WAIT_FOR, 500us
'           End If
'
'            TheExec.Datalog.WriteComment ("********************** READ @ VPP = 0 V********************************************** ")
'            With TheHdw.DCVI(g_sVPP_PINNAME)
'                 .Voltage = 5
'                 TheHdw.Wait 1 * ms
'                 .Voltage = 1.5
'                 TheHdw.Wait 1 * ms
'                 .Voltage = 0
'            End With
'            TheExec.Datalog.WriteComment ("Force " & g_sVPP_PINNAME & " = " & Format(TheHdw.DCVI.Pins(g_sVPP_PINNAME).Voltage, "0.00"))
'            'sTPPath = Application.ActiveWorkbook.Path
'            AHBScript = g_sTPPath + "\DFT\TestPlan\OTP\otp_single_prog.txt"
'            Call AHBSetup(AHBScript, NEW_PATTERN_TSU2)
'
'
'            mL_lowVal = &H5C: mL_hiVal = &H5C: AHB_READDSC "OTP_SLV_OTP_DATA0", g_RegVal
'            TestMode = "AHBOTP-Debug": SubTestCondition = "_OTP-SLV-OTP-DATA0"
'            Call UniFormTestName.CreateIDSTestName: TestName = TestMode & SubTestCondition 'UniFormTestName.Customized(Group1:=TestMode, Group2:=SubTestCondition, Group3:="X", Group4:="X")
'            Call TheExec.Flow.TestLimit(TName:=TestName, ResultVal:=g_RegVal, lowVal:=mL_lowVal, hiVal:=mL_hiVal)
'
'
'            mL_lowVal = &H7F: mL_hiVal = &H7F: AHB_READDSC "OTP_SLV_OTP_DATA1", g_RegVal
'            TestMode = "AHBOTP-Debug": SubTestCondition = "_OTP-SLV-OTP-DATA1"
'            Call UniFormTestName.CreateIDSTestName: TestName = TestMode & SubTestCondition 'UniFormTestName.Customized(Group1:=TestMode, Group2:=SubTestCondition, Group3:="X", Group4:="X")
'            Call TheExec.Flow.TestLimit(TName:=TestName, ResultVal:=g_RegVal, lowVal:=mL_lowVal, hiVal:=mL_hiVal)
'
'
'            mL_lowVal = &H9F: mL_hiVal = &H9F: AHB_READDSC "OTP_SLV_OTP_DATA2", g_RegVal
'            TestMode = "AHBOTP-Debug": SubTestCondition = "_OTP-SLV-OTP-DATA2"
'            Call UniFormTestName.CreateIDSTestName: TestName = TestMode & SubTestCondition ' UniFormTestName.Customized(Group1:=TestMode, Group2:=SubTestCondition, Group3:="X", Group4:="X")
'            Call TheExec.Flow.TestLimit(TName:=TestName, ResultVal:=g_RegVal, lowVal:=mL_lowVal, hiVal:=mL_hiVal)
'
'
'            mL_lowVal = &H1F: mL_hiVal = &H1F: AHB_READDSC "OTP_SLV_OTP_DATA3", g_RegVal
'            TestMode = "AHBOTP-Debug": SubTestCondition = "_OTP-SLV-OTP-DATA3"
'            Call UniFormTestName.CreateIDSTestName: TestName = TestMode & SubTestCondition 'UniFormTestName.Customized(Group1:=TestMode, Group2:=SubTestCondition, Group3:="X", Group4:="X")
'            Call TheExec.Flow.TestLimit(TName:=TestName, ResultVal:=g_RegVal, lowVal:=mL_lowVal, hiVal:=mL_hiVal)
'


Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function CheckAhbNOtpRead(Optional CompType As g_eAHB_OTP_COMP_TYPE = g_eAHB_OTP_COMP_TYPE.eCHECK_ALL)
    Dim sFuncName As String: sFuncName = "CheckAhbNOtpRead"
    On Error GoTo ErrHandler
    Dim lIdx As Long, lChkCnt As Long
    
    Dim aslReadOTP() As New SiteLong, aslReadAHBOTP() As New SiteLong, aslCalAhOtp() As New SiteLong
    Dim asOtpRegisterName() As String, asRegisterName() As String, asvChkResult() As New SiteVariant, asBWOffset() As String
    Dim sOtpRegisterName As String, lOtpRegOfs As Long, lBitWidth As Long
    Dim asOtpOwner() As String, asDefaultorReal() As String
    
    Select Case CompType
        Case g_eAHB_OTP_COMP_TYPE.eCHECK_BY_CONDITION
            lChkCnt = g_DictOTPPreCheckIndex.Count
        Case g_eAHB_OTP_COMP_TYPE.eCHECK_ALL
            lChkCnt = g_Total_OTP 'was UBound(g_OTPData.Category)+1 '___20200313, AHB New Method
    End Select
    
    ReDim aslReadOTP(lChkCnt - 1): ReDim aslReadAHBOTP(lChkCnt - 1): ReDim aslCalAhOtp(lChkCnt - 1)
    ReDim asOtpRegisterName(lChkCnt - 1): ReDim asRegisterName(lChkCnt - 1): ReDim asvChkResult(lChkCnt - 1): ReDim asBWOffset(lChkCnt - 1)
    ReDim asOtpOwner(lChkCnt - 1): ReDim asDefaultorReal(lChkCnt - 1) As String

    Dim lOTPCategoryIndex   As Long
    Dim lKey               As Long
    Dim lOtpIdx As Long
    
    Dim sComment           As String
    
    If TheExec.TesterMode = testModeOffline Then
       TheExec.Datalog.WriteComment ""
       TheExec.Datalog.WriteComment "** Offline Mode **"
    End If
    
    If g_bOtpEnable = False Then Exit Function
    If TheExec.Sites.Selected.Count = 0 Then Exit Function
    TheExec.Datalog.WriteComment "<FuncName> " + sFuncName + ": Please check these OTP Category Item.@ InstanceName=" + TheExec.DataManager.InstanceName
    
    '___Get All AHB Read data into OTPData whenever OTPData has the corresponding AHB register name
    Call ReadAhbToCat

    Select Case CompType
        Case g_eAHB_OTP_COMP_TYPE.eCHECK_ALL
            TheExec.Datalog.WriteComment "** <<Check All AHB-OTP>> **"
            
            For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
                With g_OTPData.Category(lOtpIdx)
                       asOtpRegisterName(lOtpIdx) = .sOtpRegisterName
                       asRegisterName(lOtpIdx) = .sRegisterName 'AHB Register Name
                       asOtpOwner(lOtpIdx) = .sOTPOwner
                       asDefaultorReal(lOtpIdx) = .sDefaultORReal
                       lBitWidth = .lBitWidth
                       lOtpRegOfs = .lOtpRegOfs
                       asBWOffset(lOtpIdx) = lBitWidth
                 End With
                '___Get OTP write(to-be-Burned) data
                aslReadOTP(lOtpIdx) = g_OTPData.Category(lOtpIdx).Read.Value
                
                '___Get AHB Read data
                Dim slTempAHBReadVal As New SiteLong
                Dim sAHBRegOtpIdx As String
                Dim asIdxArr() As String
                
                If TheExec.TesterMode = testModeOffline Then
                    If g_OTPData.Category(lOtpIdx).sAhbAddress <> "NA" Then
                        g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs = g_OTPData.Category(lOtpIdx).Read.Value
                        sAHBRegOtpIdx = g_dictAHBRegToOTPDataIdx.Item(asRegisterName(lOtpIdx)) '20190717
                        asIdxArr = Split(sAHBRegOtpIdx, ",")
                        If lOtpIdx = asIdxArr(UBound(asIdxArr)) Then
                        
                            With g_OTPData.Category(lOtpIdx)
                                For Each Site In TheExec.Sites
                                    slTempAHBReadVal(Site) = CLng(.Read.Value)
                                Next Site
                                .svAhbReadVal = slTempAHBReadVal.ShiftLeft(lOtpRegOfs)
                                If lOtpIdx <> asIdxArr(0) Then
                                    'For otpidx_of_singleAHB = 0 To UBound(asIdxArr)
                                    g_OTPData.Category(lOtpIdx).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadVal.Add(g_OTPData.Category(lOtpIdx - 1).svAhbReadVal)
                                    'update previous .svAhbReadVal
                                    g_OTPData.Category(lOtpIdx - 1).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadVal
                                    'Next otpidx_of_singleAHB
                                Else
                                    .svAhbReadVal = .svAhbReadValByMaskOfs
                                End If
                            End With
                        Else
                            g_OTPData.Category(lOtpIdx).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs
                        End If
                        aslReadAHBOTP(lOtpIdx) = g_OTPData.Category(lOtpIdx).svAhbReadVal
                        aslCalAhOtp(lOtpIdx) = g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs
                        asvChkResult(lOtpIdx) = aslCalAhOtp(lOtpIdx).Compare(EqualTo, aslReadOTP(lOtpIdx))
                    End If
                
                Else
                    If g_OTPData.Category(lOtpIdx).sAhbAddress <> "NA" Then
                        
                        aslReadAHBOTP(lOtpIdx) = g_OTPData.Category(lOtpIdx).svAhbReadVal
                        aslReadAHBOTP(lOtpIdx) = aslReadAHBOTP(lOtpIdx).ShiftRight(lOtpRegOfs)
                        '___20200313, check if Site-Loop can be removed
                        '___If YES, Set it to True
                        If (False) Then
                            aslCalAhOtp(lOtpIdx) = aslReadAHBOTP(lOtpIdx).BitWiseAnd(g_OTPData.Category(lOtpIdx).lCalDeciAhbByMaskOfs)
                        Else
                            For Each Site In TheExec.Sites
                                aslCalAhOtp(lOtpIdx)(Site) = (aslReadAHBOTP(lOtpIdx)(Site)) And (g_OTPData.Category(lOtpIdx).lCalDeciAhbByMaskOfs)
                            Next Site
                        End If
                        asvChkResult(lOtpIdx) = aslCalAhOtp(lOtpIdx).Compare(EqualTo, aslReadOTP(lOtpIdx))
                        
                    Else
                        aslReadAHBOTP(lOtpIdx) = 0
                        aslCalAhOtp(lOtpIdx) = 0
                        'mSB_Check = True
                        asvChkResult(lOtpIdx) = -999
                    End If
                End If
            Next lOtpIdx
            
        Case g_eAHB_OTP_COMP_TYPE.eCHECK_BY_CONDITION
            TheExec.Datalog.WriteComment "** <<Check AHB-OTP By Condition>> **"
            lIdx = 0
            For lKey = 1 To g_DictOTPPreCheckIndex.Count
                lOtpIdx = g_DictOTPPreCheckIndex.Item(lKey)
                'i = lOTPCategoryIndex
    
                With g_OTPData.Category(lOTPCategoryIndex)
                  sOtpRegisterName = .sOtpRegisterName
                   asOtpRegisterName(lKey - 1) = .sOtpRegisterName
                   asRegisterName(lKey - 1) = .sRegisterName
                   asOtpOwner(lKey - 1) = .sOTPOwner
                   asDefaultorReal(lKey - 1) = .sDefaultORReal
                   lBitWidth = .lBitWidth
                   lOtpRegOfs = .lOtpRegOfs
                   'asBWOffset(lIdx) = "[" & FormatLog((lBitWidth + lOtpRegOfs - 1), 1) & ":" & FormatLog(lOtpRegOfs, 1) & "]"
                   asBWOffset(lKey - 1) = lBitWidth
               End With
                '___Get OTP write(to-be-Burned) data
                aslReadOTP(lKey - 1) = g_OTPData.Category(lOtpIdx).Read.Value
                
                If TheExec.TesterMode = testModeOffline Then
                    If g_OTPData.Category(lOtpIdx).sAhbAddress <> "NA" Then
                        g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs = g_OTPData.Category(lOtpIdx).Read.Value
                        sAHBRegOtpIdx = g_dictAHBRegToOTPDataIdx.Item(asRegisterName(lKey - 1)) '20190717
                        asIdxArr = Split(sAHBRegOtpIdx, ",")
                        If lOtpIdx = asIdxArr(UBound(asIdxArr)) Then
                        
                            With g_OTPData.Category(lOtpIdx)
                                '___20200313, check if Site-Loop can be removed
                                '___If YES, Set it to True
                                If (False) Then
                                    slTempAHBReadVal = .Read.Value
                                Else
                                    For Each Site In TheExec.Sites
                                        slTempAHBReadVal(Site) = CLng(.Read.Value)
                                    Next Site
                                End If
                                .svAhbReadVal = slTempAHBReadVal.ShiftLeft(lOtpRegOfs)
                                If lOtpIdx <> asIdxArr(0) Then
                                    'For otpidx_of_singleAHB = 0 To UBound(asIdxArr)
                                    g_OTPData.Category(lOtpIdx).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadVal.Add(g_OTPData.Category(lOtpIdx - 1).svAhbReadVal)
                                    'update previous .svAhbReadVal
                                    g_OTPData.Category(lOtpIdx - 1).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadVal
                                    'Next otpidx_of_singleAHB
                                Else
                                    .svAhbReadVal = .svAhbReadValByMaskOfs
                                End If
                            End With
                        Else
                            g_OTPData.Category(lOtpIdx).svAhbReadVal = g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs
                        End If
                        aslReadAHBOTP(lKey - 1) = g_OTPData.Category(lOtpIdx).svAhbReadVal
                        aslCalAhOtp(lKey - 1) = g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs
                        asvChkResult(lKey - 1) = aslCalAhOtp(lKey - 1).Compare(EqualTo, aslReadOTP(lKey - 1))
                        If asvChkResult(lKey - 1)(0) Like "*False*" Then Stop
                    End If
                    
                Else
                    '1).Get OTP Read data
                    Call auto_OTPCategory_GetReadDecimal(sOtpRegisterName, aslReadOTP(lKey - 1)) ', False)
                    '2).Get AHB Read data
                    Call auto_OTPCategory_GetReadDecimal_AHB(sOtpRegisterName, aslReadAHBOTP(lKey - 1)) ', False)
                    
                    If TheExec.TesterMode = testModeOnline Then
                       aslCalAhOtp(lKey - 1) = aslReadAHBOTP(lKey - 1)
                    Else
                        aslCalAhOtp(lKey - 1) = aslReadOTP(lKey - 1)
                    End If
                    asvChkResult(lKey - 1) = aslCalAhOtp(lKey - 1).Compare(EqualTo, aslReadOTP(lKey - 1))
                End If
            Next lKey
        End Select
     'B).DataLog
        Dim sTemp As String
        Dim slFailCnt As New SiteLong

        TheExec.Datalog.WriteComment "<" + sFuncName + ">, InstanceName::" & TheExec.DataManager.InstanceName
        TheExec.Datalog.WriteComment "<" + sFuncName + ">, AHB CHECK CNT       :" & lChkCnt
        
        If CompType = g_eAHB_OTP_COMP_TYPE.eCHECK_ALL Then
           TheExec.Datalog.WriteComment "<" + sFuncName + ">, CHECK All OTP-AHB"
        Else
           TheExec.Datalog.WriteComment "<" + sFuncName + ">, CHECK OTPOwner      :" & g_sOTP_OWNER_FOR_CHECK
           TheExec.Datalog.WriteComment "<" + sFuncName + ">, CHECK DefaultORReal :" & g_sOTP_DEFAULT_REAL_FOR_PRECHECK
        End If
        TheExec.Datalog.WriteComment "<" + sFuncName + ">, PLEASE CHECK THESE OTP/AHB LISTS!"


        For Each Site In TheExec.Sites.Selected
            sTemp = "" 'FormatLog("Comment", -20)
            sTemp = " [Site" & CStr(Site) & "]"
            sTemp = sTemp & "," & FormatLog("OTP-REGName", -g_lOTPCateNameMaxLen - 17) & "," & FormatLog("OTP-TrimCode", -15)
            sTemp = sTemp & "," & FormatLog("AHB-REGName", -g_lOTPCateNameMaxLen) & "," & FormatLog("AHB-TrimCode", -15)
            sTemp = sTemp & "," & FormatLog("Check", -10)
            sTemp = sTemp & "," & FormatLog("BW", -5)
            sTemp = sTemp & "," & FormatLog("OTPOWNER", -10)
            sTemp = sTemp & "," & FormatLog("Default|Real", -15)

            TheExec.Datalog.WriteComment ""
            TheExec.Datalog.WriteComment "*********************************************************************************************** " & _
                                       "*********************************************************************************************************"
            TheExec.Datalog.WriteComment sTemp
            TheExec.Datalog.WriteComment "*********************************************************************************************** " & _
                                       "*********************************************************************************************************"
            slFailCnt(Site) = 0
            For lIdx = 0 To UBound(asOtpRegisterName)
                If InStr(UCase(g_asBYPASS_AHBOTP_CHECK), UCase(asOtpRegisterName(lIdx))) = 0 Then
                    sComment = ""
                Else
                    sComment = "" '2018-12-20: Need To Check All.
                    'sComment = "(NoNeedToCheckOTP)"
                End If
                
                sTemp = ""
                sTemp = " [Site" & CStr(Site) & "]"
                'sTemp = sTemp & "," & FormatLog(asOtpRegisterName(lIdx), -g_lOTPCateNameMaxLen) & "," & FormatLog("&H" + Hex(aslReadOTP(lIdx)), -12)
                'sTemp = sTemp & "," & FormatLog(asRegisterName(lIdx), -g_lOTPCateNameMaxLen - 10) & "," & FormatLog("&H" + Hex(aslReadAHBOTP(lIdx)), -12)
                sTemp = sTemp & "," & FormatLog(asOtpRegisterName(lIdx), -g_lOTPCateNameMaxLen) & "," & FormatLog(sComment, -17) & "," & FormatLog(aslReadOTP(lIdx), -15)
                sTemp = sTemp & "," & FormatLog(asRegisterName(lIdx), -g_lOTPCateNameMaxLen) & "," & FormatLog(aslCalAhOtp(lIdx), -15)
                sTemp = sTemp & "," & FormatLog(asvChkResult(lIdx), -10)
                sTemp = sTemp & "," & FormatLog(asBWOffset(lIdx), -5)
                sTemp = sTemp & "," & FormatLog(asOtpOwner(lIdx), -10)
                sTemp = sTemp & "," & FormatLog(asDefaultorReal(lIdx), -15)
                
                '2018-12-20: Need To Check All.
                If asvChkResult(lIdx) <> -1 Then slFailCnt(Site) = slFailCnt(Site) + 1
                'If asvChkResult(lIdx) <> 0 And (InStr(UCase(g_asBYPASS_AHBOTP_CHECK), UCase(asOtpRegisterName(lIdx))) = 0) Then slFailCnt(Site) = slFailCnt(Site) + 1
                
                'Datalog for check fail only
                If asvChkResult(lIdx) <> -1 Then
                    TheExec.Datalog.WriteComment sTemp
                End If
            Next lIdx
            TheExec.Datalog.WriteComment ""
        Next Site

        'DATALOG:
        Dim ParmName As String
        ParmName = "AHBOTPReadPostCheck-FailCNT"

        '___User Maintain high Limit
        TheExec.Flow.TestLimit slFailCnt, 0, 0, TName:=ParmName



Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function BurnWriteOTP_FW(r_psFWTsuDNGS As PatternSet, r_psFWTsuSYS As PatternSet, r_psFWTsuCRC As PatternSet)
    Dim sFuncName As String: sFuncName = "BurnWriteOTP_FW"
    On Error GoTo ErrHandler
    Dim sdMonitorVolt As New SiteDouble
    Dim sbPFStatus            As New SiteBoolean
    Dim lFailCnt          As Long
    '___offline Debug without patterns on MP4T trunk
    If TheExec.TesterMode = testModeOffline Then
        r_psFWTsuDNGS.Value = ".\Patterns\CPU\JTAG\PP_SUZA0_C_IN00_JT_XXXX_WIR_JTG_UNS_ALLFV_SI_AHB_SEL_1_A0_1810050226.PAT" '".\Patterns\CPU\ANALOGUE\DD_AVSA0_C_IN00_AN_XXXX_PFF_JTG_UNS_ALLFV_SI_DESIGNRG_APZFWTSU_2_A0_1901231059.PAT"
        r_psFWTsuSYS.Value = ".\Patterns\CPU\JTAG\PP_SUZA0_C_IN00_JT_XXXX_WIR_JTG_UNS_ALLFV_SI_AHB_SEL_1_A0_1810050226.PAT" ' ".\Patterns\CPU\ANALOGUE\DD_AVSA0_C_IN00_AN_XXXX_PFF_JTG_UNS_ALLFV_SI_SYSTEMRG_APZFWTSU_2_A0_1901231059.PAT"
        r_psFWTsuCRC.Value = ".\Patterns\CPU\JTAG\PP_SUZA0_C_IN00_JT_XXXX_WIR_JTG_UNS_ALLFV_SI_AHB_SEL_1_A0_1810050226.PAT"
    End If
    TheHdw.Patterns(r_psFWTsuDNGS).Load
    TheHdw.Patterns(r_psFWTsuSYS).Load
    
    Call ForceVbyPPMU("GPIO18", 1.8)
    
    TheHdw.StartStopwatch 'Timer start
    TheHdw.Patterns(r_psFWTsuDNGS).Start: TheHdw.Digital.Patgen.HaltWait
    TheExec.Datalog.WriteComment "Running DESIGN Pattern:" & r_psFWTsuDNGS.Value
    TheHdw.Patterns(r_psFWTsuSYS).Start: TheHdw.Digital.Patgen.HaltWait
    TheExec.Datalog.WriteComment "Running SYSTEM Pattern:" & r_psFWTsuSYS.Value
    Call OTP_SPT_D(" *** OTP DESIGN/SYSTEM  , Exe Time =  ") 'Timer stop
    

    TheHdw.StartStopwatch 'Timer start
    TheHdw.Patterns(r_psFWTsuCRC).Load
    TheHdw.Patterns(r_psFWTsuCRC).Start: TheHdw.Digital.Patgen.HaltWait
    TheExec.Datalog.WriteComment "Running OTP FW Pattern:" & r_psFWTsuCRC.Value
    Call OTP_SPT_D(" *** OTP FW Pattern , Exe Time =  ") 'Timer stop
        
    '//Check Pattern Burst Pass/Fail
    sbPFStatus = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    TheExec.Flow.TestLimit sbPFStatus, -1, 1, , , , unitNone, , TName:="OTP_FWPattern_Check"
    If (0#) Then
        For Each g_Site In TheExec.Sites
            lFailCnt = TheHdw.Digital.FailedPinsCount(g_Site)
            If lFailCnt = 0 Then
               TheExec.Datalog.WriteComment "Site " & CStr(g_Site) & " PASS - " & r_psFWTsuCRC.Value
            Else
                TheExec.Datalog.WriteComment "Site " & CStr(g_Site) & " FAIL - " & r_psFWTsuCRC.Value
            End If
        Next g_Site
    End If
    TheHdw.StartStopwatch
    sdMonitorVolt = MeasureVbyPPMU("GPIO19")
    Call TheExec.Flow.TestLimit(sdMonitorVolt, 1, , , , scaleNone, unitVolt, "%6.4f", "OTP_Done_GPIO19_Check", , "NA", , unitCustom, unitVolt, unitVolt, tlForceNone)
    Call OTP_SPT_D(" *** OTP Monitor Voltage , Exe Time =  ")
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function LoadAhbFromOtp()
    Dim sFuncName As String: sFuncName = "LoadAhbFromOtp"
    On Error GoTo ErrHandler
    Dim bSupress As Boolean: bSupress = True
    Dim sBuck1PFBPins As String
    Dim sVddBuckPins As String
    Dim sbSaveSiteStatus As New SiteBoolean

    sbSaveSiteStatus = TheExec.Sites.Selected

    SetNwireEnableFlag True
    SetSpmiEnableFlag True

    ''20200220
        '' ----------------------------------------------switch to Nwire SPMI----------------------------------------------------
        TheExec.Sites.Selected = TheExec.Sites.Existing
        SetBitFieldEnableFlag True
        '            RegVal = &H10: AHB_WRITEDSC HOST_INTERFACE_SPMI_REGISTERS_SPMI_CTRL.Addr, RegVal
        g_RegVal = &H10: AHB_WRITEDSC "HOST_INTERFACE_SPMI_REGISTERS_SPMI_CTRL", g_RegVal
        
        TheHdw.Protocol.ports("NWIRE_SPMI").Enabled = True
        TheHdw.Protocol.ports("nWire_SPMI").NWire.HRAM.Setup.TriggerType = tlNWireHRAMTriggerType_Never
        TheHdw.Protocol.ports("nWire_SPMI").NWire.HRAM.Setup.WaitForEvent = False
        TheHdw.Protocol.ModuleRecordingEnabled = True '?
        'Call ENABLE_SPMI_PA
        
        SetNwireEnableFlag True
        SetSpmiEnableFlag True


        g_RegVal = &H1: AHB_WRITE "FABRIC_AHB_FABRIC_RMWU_MAIN_EN", g_RegVal
        g_RegVal = &H2: AHB_WRITE "FABRIC_AHB_FABRIC_RMWU_MODE", g_RegVal
        
        TheExec.Datalog.WriteComment "SWITCH to SPMI PA!!!"
        
            sbSaveSiteStatus = TheExec.Sites.Selected
        ' ----------------------------------------------End switch to Nwire SPMI------------------------------------------------
    
    'Only reload burned die / frsh die will keep trim section AHB Value
    If g_sbOtpedPGM.All(False) Then Exit Function

bSupress = True

    sBuck1PFBPins = "BUCK3_FB_UVI80,BUCK9_FB_UVI80,BUCK14_FB_UVI80"
    sVddBuckPins = "VDD_BUCK3_14_UVI80,VDD_BUCK1_8_9_UVI80"
    
    '___20200313, Copy from MP7P, it could be different here by projects
    If (False) Then
        '___20200313, New AHB Method
        If (bSupress) Then
            g_RegVal = &HF: AHB_WRITE "TST_CTRL_DFT_FORCE", g_RegVal    'enter testmode
            g_RegVal = &H1: AHB_WRITE "DVC_SCHEDULER_DVC_TEST_ALLOW_VSEL_WRITE_WHEN_DISABLED", g_RegVal
        End If
    
        With TheHdw.DCVI.Pins(sBuck1PFBPins)
            .Gate = False
            .Connect
            .CurrentRange = 0.2
            .Current = 0.02
            .Voltage = 1.5 'g_VDD_1p5V_VDDANA
            .Gate = True
        End With
        
        TheHdw.DCVI.Pins(sVddBuckPins).Voltage = 2.5
        
        TheHdw.Utility.Pins("K1460,K1461,K1560,K1561,K1660,K1661").State = tlUtilBitOn

        TheExec.Sites.Selected = g_sbOtpedPGM
'    If gB_OTP_Enable = True Then
        ''1) RUN pattern IN00
        TheHdw.Patterns(".\Patterns\CPU\ANALOGUE\PP_SERA0_C_IN00_AN_XXXX_PFF_JTG_UNS_ALLFRV_SI_RELOAD_OTPTSU_1_A0_1909032218.PAT.gz").Load
        TheHdw.Patterns(".\Patterns\CPU\ANALOGUE\PP_SERA0_C_IN00_AN_XXXX_PFF_JTG_UNS_ALLFRV_SI_RELOAD_OTPTSU_1_A0_1909032218.PAT.gz").Start
        TheHdw.Digital.Patgen.HaltWait

        'Power VSS_DFT_2 to 1.5
        TheHdw.Digital.Pins("VSS_DFT_2").Disconnect
        With TheHdw.PPMU.Pins("VSS_DFT_2")
            .Connect
            .Gate = tlOn
            .ForceV 1.5
            TheHdw.Wait 600 * ns
        End With

        ''3) RUN pattern IN01
        TheHdw.Patterns(".\Patterns\CPU\ANALOGUE\PP_SERA0_C_IN01_AN_XXXX_PFF_JTG_UNS_ALLFRV_SI_RELOAD_OTPTSU_1_A0_1909032218.PAT.gz").Load
        TheHdw.Patterns(".\Patterns\CPU\ANALOGUE\PP_SERA0_C_IN01_AN_XXXX_PFF_JTG_UNS_ALLFRV_SI_RELOAD_OTPTSU_1_A0_1909032218.PAT.gz").Start
        TheHdw.Digital.Patgen.HaltWait

        ''4) RUN pattern IN02
        TheHdw.Patterns(".\Patterns\CPU\ANALOGUE\PP_SERA0_C_IN02_AN_XXXX_PFF_JTG_UNS_ALLFRV_SI_RELOAD_OTPTSU_1_A0_1909032218.PAT.gz").Load
        TheHdw.Patterns(".\Patterns\CPU\ANALOGUE\PP_SERA0_C_IN02_AN_XXXX_PFF_JTG_UNS_ALLFRV_SI_RELOAD_OTPTSU_1_A0_1909032218.PAT.gz").Start
        TheHdw.Digital.Patgen.HaltWait
        
        '___After the OTPTSU, the SPMI must switch again.
        SetNwireEnableFlag True
        SetSpmiEnableFlag True
        '' ----------------------------------------------switch to Nwire SPMI----------------------------------------------------
        TheExec.Sites.Selected = TheExec.Sites.Existing
            SetBitFieldEnableFlag True
'            g_RegVal = &H10: AHB_WRITEDSC HOST_INTERFACE_SPMI_REGISTERS_SPMI_CTRL.Addr, g_RegVal
            g_RegVal = &H10: AHB_WRITEDSC "HOST_INTERFACE_SPMI_REGISTERS_SPMI_CTRL", g_RegVal
        
            TheHdw.Protocol.ports("NWIRE_SPMI").Enabled = True
            TheHdw.Protocol.ports("nWire_SPMI").NWire.HRAM.Setup.TriggerType = tlNWireHRAMTriggerType_Never
            TheHdw.Protocol.ports("nWire_SPMI").NWire.HRAM.Setup.WaitForEvent = False
            TheHdw.Protocol.ModuleRecordingEnabled = True '?
            'Call ENABLE_SPMI_PA
        
            SetNwireEnableFlag True
            SetSpmiEnableFlag True
        
            g_RegVal = &H1: AHB_WRITE "FABRIC_AHB_FABRIC_RMWU_MAIN_EN", g_RegVal
            g_RegVal = &H2: AHB_WRITE "FABRIC_AHB_FABRIC_RMWU_MODE", g_RegVal
        
            TheExec.Datalog.WriteComment "SWITCH to SPMI PA!!!"
        TheExec.Sites.Selected = sbSaveSiteStatus
        ' ----------------------------------------------End switch to Nwire SPMI------------------------------------------------

        If GetModBurstStatus = False Then
            g_RegVal = &H61: AHB_WRITE "SEC_SYSCTRL_TEST_REG_EN0", g_RegVal
            g_RegVal = &H45: AHB_WRITE "SEC_SYSCTRL_TEST_REG_EN1", g_RegVal
            g_RegVal = &H72: AHB_WRITE "SEC_SYSCTRL_TEST_REG_EN2", g_RegVal
            g_RegVal = &H4F: AHB_WRITE "SEC_SYSCTRL_TEST_REG_EN3", g_RegVal
            g_RegVal = &H1: AHB_WRITE "POWER_CONTROL_MAINFSM_OTP_DFT_ARCH_0.SPMI_DEBUG_EN", g_RegVal
            ' Override pattern setting
            
            '<Notice>
            'Call OTP_PostBurn_AHB_SetUp ''''this one is different by projects
            
        ElseIf TheHdw.Protocol.ports("NWIRE_SPMI").Modules.IsRecorded("OTP_PostBurn_AHB_SetUp") = False Then 'If module burst is enabled and has been recorded, do module burst directly
            g_RegVal = &H61: AHB_WRITE "SEC_SYSCTRL_TEST_REG_EN0", g_RegVal
            g_RegVal = &H45: AHB_WRITE "SEC_SYSCTRL_TEST_REG_EN1", g_RegVal
            g_RegVal = &H72: AHB_WRITE "SEC_SYSCTRL_TEST_REG_EN2", g_RegVal
            g_RegVal = &H4F: AHB_WRITE "SEC_SYSCTRL_TEST_REG_EN3", g_RegVal
            g_RegVal = &H1: AHB_WRITE "POWER_CONTROL_MAINFSM_OTP_DFT_ARCH_0.SPMI_DEBUG_EN", g_RegVal
            ' Override pattern setting
            
            '<Notice>
            'Call OTP_PostBurn_AHB_SetUp ''''this one is different by projects
            
            Call TheHdw.Protocol.ports("NWIRE_SPMI").Modules.StopRecording
        End If
'    End If
        TheExec.Sites.Selected = sbSaveSiteStatus
    
        TheHdw.Utility.Pins("K1460,K1461,K1560,K1561,K1660,K1661").State = tlUtilBitOff
        TheHdw.DCVI.Pins(sVddBuckPins).Voltage = 0 'g_VDD_BUCK_UVI
        
        With TheHdw.DCVI.Pins(sBuck1PFBPins)
            .Gate = False
            .Disconnect
        End With
    
        If (bSupress) Then
                g_RegVal = &H0: AHB_WRITE "TST_CTRL_DFT_FORCE", g_RegVal    'enter testmode
                g_RegVal = &H0: AHB_WRITE "DVC_SCHEDULER_DVC_TEST_ALLOW_VSEL_WRITE_WHEN_DISABLED", g_RegVal
        End If
    End If
    '******************************************User Maintain******************************************

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ActivateFSM(AHB_SEL_Pat As PatternSet)
    Dim sFuncName As String: sFuncName = "ActivateFSM"
    On Error GoTo ErrHandler
    
    'Debug.Print "User maintains this function"
    '******************************************User Maintain******************************************
        'Enable AHB
        'TheHdw.PPMU.Pins(g_sDIG_PINS).Disconnect
        'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        'GetPatListNExecutePat AHB_SEL_Pat
        'Enable ATB
        'GetPatListNExecutePat "ATB_EN_TSU00"
    
    '    For Each g_Site In TheExec.Sites
    '         If TheExec.TesterMode = testModeOffline Then g_RegVal = &H91
    '         TheExec.Datalog.WriteComment ("*** Excute FSM ACTIVE*** Site " & g_Site & "  Read IDCODE= &H" & Hex(g_RegVal(g_Site)))
    '    Next g_Site
    '******************************************User Maintain******************************************
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ForcePOR(ForcePOR_Pat As PatternSet)
    On Error GoTo ErrHandler
    Dim sFuncName As String: sFuncName = "ForcePOR"
    
    '******************************************User Maintain******************************************
    'Force Power Reset
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    GetPatListNExecutePat ForcePOR_Pat
    '******************************************User Maintain******************************************

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20190723 update
'___Print the expect/actual/compared OTP results on Datalog
Public Function CompareExpActDatalog(r_lAddrStart As Long, r_lAddrEnd As Long, _
                                   r_wPGMData As DSPWave, r_wReadData As DSPWave, r_wPGMCompreRead As DSPWave) As Long
    Dim sFuncName As String: sFuncName = "CompareExpActDatalog"
    On Error GoTo ErrHandler
    Dim lAddrIdx As Long
    Dim adPGMArr() As Double
    Dim adReadArr() As Double
    ReDim g_asLogTestName(11)
    
    '___TestName
    g_asLogTestName(0) = "OTP"
    g_asLogTestName(1) = "Read"  ' SubTestMode
    g_asLogTestName(2) = "0d"
    g_asLogTestName(3) = "X"
    g_asLogTestName(4) = "X"
    g_asLogTestName(5) = "OTPOffset-&H" & Format(Hex(g_iOTP_ADDR_OFFSET), "000")  ' "X"
    g_asLogTestName(6) = "Addr"
    ''g_asLogTestName(7) = Right("000" & CStr(r_lAddrStart), 3) 'group 8, to always be 3 digits
    g_asLogTestName(8) = Replace(m_VddLevel, "_", "")
    g_asLogTestName(9) = "Expected"
    g_asLogTestName(10) = "X"
    g_asLogTestName(11) = "X"
    
    
    For lAddrIdx = r_lAddrStart To r_lAddrEnd
        g_asLogTestName(7) = Right("0000" & CStr(lAddrIdx), 4) 'change to 4 digits 'group 8, to always be 3 digits
    
        '___Step1:
        For Each Site In TheExec.Sites
            adPGMArr = r_wPGMData.Data
            adReadArr = r_wReadData.Data
            
            TheExec.Datalog.WriteComment ""
               
        '___Step2:Log the Program Data & Expected Data:
            m_TestName = TNameCombine(g_asLogTestName(0), g_asLogTestName(1), "0d", g_asLogTestName(3), g_asLogTestName(4), g_asLogTestName(5), TName_OTP_Addr, g_asLogTestName(7), g_asLogTestName(8), TName_OTP_Grp10_Expected)
            Call TheExec.Flow.TestLimit(TName:=m_TestName, ResultVal:=adPGMArr(lAddrIdx), lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, formatStr:="%.0f", scaletype:=scaleNoScaling) ' formatStr:="%.6f") ' lowval:=0, hival:=255
            'Call TheExec.Flow.TestLimit(TName:=m_TestName, ResultVal:=PGM_dataWave(Site).Element(lAddrIdx), lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, formatStr:="%.0f", scaletype:=scaleNoScaling)
             
             
            m_TestName = TNameCombine(g_asLogTestName(0), g_asLogTestName(1), "0d", g_asLogTestName(3), g_asLogTestName(4), g_asLogTestName(5), _
                     TName_OTP_Addr, g_asLogTestName(7), g_asLogTestName(8), TName_OTP_Grp10_Actual)
            'TestName = TestBlock & TestMode & SubTestMode & SubTestCondition & m_VddLevel & MeasureType & "_Addr" & lAddrIdx & "_OTP_actual_0d"
            Call TheExec.Flow.TestLimit(TName:=m_TestName, ResultVal:=adReadArr(lAddrIdx), lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, formatStr:="%.0f", scaletype:=scaleNoScaling) ' formatStr:="%.6f") ' lowval:=0, hival:=255
            
            
         '___Step3:Compare the Program Data & Expected Data:
            If (g_sbOtpedPGM(Site) = False) Then
    
                m_TestName = TNameCombine(g_asLogTestName(0), g_asLogTestName(1), "X", g_asLogTestName(3), g_asLogTestName(4), g_asLogTestName(5), _
                TName_OTP_Addr, g_asLogTestName(7), g_asLogTestName(8), TName_OTP_Grp10_Match)
                If lAddrIdx < 4 Then
                    Call TheExec.Flow.TestLimit(TName:=m_TestName, ResultVal:=r_wPGMCompreRead(Site).ElementLite(lAddrIdx), lowVal:=-1, hiVal:=-1, lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, formatStr:="%.0f")
                Else
                    Call TheExec.Flow.TestLimit(TName:=m_TestName, ResultVal:=r_wPGMCompreRead(Site).ElementLite(lAddrIdx), lowVal:=-1, hiVal:=-1, lowCompareSign:=tlSignEqual, highCompareSign:=tlSignEqual, formatStr:="%.0f")
                End If
            Else
             
                m_TestName = TNameCombine(g_asLogTestName(0), g_asLogTestName(1), "X", g_asLogTestName(3), g_asLogTestName(4), g_asLogTestName(5), TName_OTP_Addr, g_asLogTestName(7), g_asLogTestName(8), TName_OTP_Grp10_Match)
                Call TheExec.Flow.TestLimit(TName:=m_TestName, ResultVal:=r_wPGMCompreRead(Site).ElementLite(lAddrIdx), lowVal:=0, hiVal:=0, lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, formatStr:="%.0f") ' formatStr:="%.6f")
             End If
        Next Site
    Next lAddrIdx

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CompareExpActDatalog_read_only(r_lAddrStart As Long, r_lAddrEnd As Long, r_wReadData As DSPWave) As Long
    Dim sFuncName As String: sFuncName = "CompareExpActDatalog_read_only"
    On Error GoTo ErrHandler
    Dim lAddrIdx As Long
    Dim adReadArr() As Double
    ReDim g_asLogTestName(11)
    
    '___TestName
    g_asLogTestName(0) = "OTP"
    g_asLogTestName(1) = "Read"  ' SubTestMode
    g_asLogTestName(2) = "0d"
    g_asLogTestName(3) = "X"
    g_asLogTestName(4) = "X"
    g_asLogTestName(5) = "OTPOffset-&H" & Format(Hex(g_iOTP_ADDR_OFFSET), "000")  ' "X"
    g_asLogTestName(6) = "Addr"
    ''g_asLogTestName(7) = Right("000" & CStr(r_lAddrStart), 3) 'group 8, to always be 3 digits
    g_asLogTestName(8) = Replace(m_VddLevel, "_", "")
    g_asLogTestName(9) = "Expected"
    g_asLogTestName(10) = "X"
    g_asLogTestName(11) = "X"

    For lAddrIdx = r_lAddrStart To r_lAddrEnd
        g_asLogTestName(7) = Right("0000" & CStr(lAddrIdx), 4) 'change to 4 digits 'group 8, to always be 3 digits
        For Each Site In TheExec.Sites
            adReadArr = r_wReadData.Data
            TheExec.Datalog.WriteComment ""
            g_sSubTestCondition = "_RDbackValue"
        '___:Log the Programmed Data:
            m_TestName = TNameCombine(g_asLogTestName(0), g_asLogTestName(1), "0d", g_asLogTestName(3), g_asLogTestName(4), g_asLogTestName(5), _
                     TName_OTP_Addr, g_asLogTestName(7), g_asLogTestName(8), TName_OTP_Grp10_Actual)
            'TestName = TestBlock & TestMode & SubTestMode & SubTestCondition & m_VddLevel & MeasureType & "_Addr" & lAddrIdx & "_OTP_actual_0d"
            Call TheExec.Flow.TestLimit(TName:=m_TestName, ResultVal:=adReadArr(lAddrIdx), lowCompareSign:=tlSignNone, highCompareSign:=tlSignNone, formatStr:="%.0f", scaletype:=scaleNoScaling) ' formatStr:="%.6f") ' lowval:=0, hival:=255
        Next Site
    Next lAddrIdx

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function SimulateTrimSection()
    Dim sFuncName As String: sFuncName = "SimulateTrimSection"
    On Error GoTo ErrHandler
    If TheExec.TesterMode = testModeOffline Then
        Dim slWriteValue As New SiteLong
        Dim lOtpIdx As Long
    
    ''' --------- set fake value ---------
        For Each Site In TheExec.Sites
            slWriteValue(Site) = 1
        Next Site
        
    '    For lOtpIdx = 13 To UBound(g_OTPData.Category) - 1
    '        If LCase(g_OTPData.Category(lOtpIdx).sDefaultorReal) Like "*real*" Or _
    '        ((UCase(g_OTPData.Category(lOtpIdx).sOtpRegisterName)) Like "*OTP_WLED_CFG_IDAC_BIT0_THERMO_TRIM_1117*") Then
    '            Call auto_OTPCategory_SetWriteDecimal(g_OTPData.Category(lOtpIdx).sOtpRegisterName, slWriteValue)
    '        End If
    '    Next lOtpIdx
    '
        For lOtpIdx = 13 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
            If LCase(g_OTPData.Category(lOtpIdx).sDefaultORReal) Like "*real*" Then
                If UCase(g_OTPData.Category(lOtpIdx).sOtpRegisterName) Like "*OTP_OTP_SLV_MINOR_OTP_VERSION_4432*" _
                    Or UCase(g_OTPData.Category(lOtpIdx).sOtpRegisterName) Like "*OTP_OTP_SLV_MAJOR_OTP_VERSION_4432*" _
                    Or UCase(g_OTPData.Category(lOtpIdx).sOtpRegisterName) Like "*OTP_OTP_SLV_LCK0_4433*" _
                    Or UCase(g_OTPData.Category(lOtpIdx).sOtpRegisterName) Like "*OTP_OTP_SLV_LCK1_4434*" _
                    Or UCase(g_OTPData.Category(lOtpIdx).sOtpRegisterName) Like "*OTP__CRC_0*" Then
                    Debug.Print "Ignore " & g_OTPData.Category(lOtpIdx).sOtpRegisterName & " Offline SetWrite"
                Else
                    Call auto_OTPCategory_SetWriteDecimal(g_OTPData.Category(lOtpIdx).sOtpRegisterName, slWriteValue)
                End If
            End If
        Next lOtpIdx
    ''' --------- set fake value ---------
    End If
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function LoadAhbDefaultValue_basedOnOtpVersion()
    Dim sFuncName As String: sFuncName = "LoadAhbDefaultValue_basedOnOtpVersion"
    On Error GoTo ErrHandler
    Dim sBuck1PFBPins As String
    Dim sVddBuckPins As String
    Dim bSupress As Boolean
    Dim sbSaveSiteStatus As New SiteBoolean
    Dim sbRunDefPat As New SiteBoolean

    If g_sOTPRevisionType = "OTP_V01" Then Exit Function

'    bSupress = True
'    sbSaveSiteStatus = TheExec.Sites.Selected
'
'    For Each Site In TheExec.Sites.Selected '.Existing
'        If TheExec.Sites(Site).SiteVariableValue("RunTrim") = -1 Then
'            sbRunDefPat(Site) = True
'        End If
'    Next Site
'    TheExec.Sites.Selected.Value = sbRunDefPat
'    If sbRunDefPat.All(False) = True Then GoTo Skip_DefPat
'
''----------------------------------Only Selected Site execute the default pattern---------------------------------------------------------
'    sBuck1PFBPins = "BUCK3_FB_UVI80,BUCK9_FB_UVI80,BUCK14_FB_UVI80"
'    sVddBuckPins = "VDD_BUCK3_14_UVI80,VDD_BUCK1_8_9_UVI80"
'
'If (bSupress) Then
'        g_RegVal = &HF: AHB_WRITE "TST_CTRL_DFT_FORCE", g_RegVal    'enter testmode
'        g_RegVal = &H1: AHB_WRITE "DVC_SCHEDULER_DVC_TEST_ALLOW_VSEL_WRITE_WHEN_DISABLED", g_RegVal
'
'End If
'
'
''''''ATE Safety net
'
'        With TheHdw.DCVI.Pins(sBuck1PFBPins)
'            .Gate = False
'            .Connect
'            .CurrentRange = 0.2
'            .Current = 0.02
'            .Voltage = g_VDD_1p5V_VDDANA
'            .Gate = True
'        End With
'
'        TheHdw.DCVI.Pins(sVddBuckPins).Voltage = 2.5
'
'        TheHdw.Utility.Pins("K1460,K1461,K1560,K1561,K1660,K1661").State = tlUtilBitOn
'
        ' JY 20200318 Run Default pattern
            TheExec.Datalog.WriteComment String(100, "-") & Chr(10) & "Running Default Pattern : " & g_sOTPRevisionType & Chr(10) & String(100, "-")
            TheHdw.Patterns(g_sOTPRevisionType).Load
            TheHdw.Patterns(g_sOTPRevisionType).Start
            TheHdw.Digital.Patgen.HaltWait
'
'        g_RegVal = &H1: AHB_WRITE "POWER_CONTROL_MAINFSM_OTP_DFT_ARCH_0.SPMI_DEBUG_EN", g_RegVal
'
'''20200220
''' ----------------------------------------------switch to Nwire SPMI----------------------------------------------------
'TheExec.Sites.Selected = TheExec.Sites.Existing
'    SetBitFieldEnableFlag True
'    g_RegVal = &H10: AHB_WRITEDSC "HOST_INTERFACE_SPMI_REGISTERS_SPMI_CTRL", g_RegVal
'
'    TheHdw.Protocol.ports("NWIRE_SPMI").Enabled = True
'    TheHdw.Protocol.ports("nWire_SPMI").NWire.HRAM.Setup.TriggerType = tlNWireHRAMTriggerType_Never
'    TheHdw.Protocol.ports("nWire_SPMI").NWire.HRAM.Setup.WaitForEvent = False
'    TheHdw.Protocol.ModuleRecordingEnabled = True '?
'    'Call ENABLE_SPMI_PA
'
'    SetNwireEnableFlag True
'    SetSpmiEnableFlag True
'
'
'    g_RegVal = &H1: AHB_WRITE "FABRIC_AHB_FABRIC_RMWU_MAIN_EN", g_RegVal
'    g_RegVal = &H2: AHB_WRITE "FABRIC_AHB_FABRIC_RMWU_MODE", g_RegVal
'
'    TheExec.Datalog.WriteComment "SWITCH to SPMI PA!!!"
'TheExec.Sites.Selected = sbRunDefPat
'' ----------------------------------------------End switch to Nwire SPMI------------------------------------------------
'
''2020-02-05
''BSTLQ Register setting experiment from Moises
'AHB_READ "BSTLQ_DIG_LOCAL_1", g_RegVal
'g_RegVal = 2: AHB_WRITE "BSTLQ_DIG_LOCAL_1", g_RegVal
'
'
''
''   Override pattern setting
''___20200220 Add Module Burst
'    If GetModBurstStatus = False Then
'       Call OTP_PostBurn_AHB_SetUp
'    ElseIf TheHdw.Protocol.ports("NWIRE_SPMI").Modules.IsRecorded("OTP_DefPatt_AHB_SetUp") = False Then 'If module burst is enabled and has been recorded, do module burst directly
'       Call OTP_PostBurn_AHB_SetUp
'       Call TheHdw.Protocol.ports("NWIRE_SPMI").Modules.StopRecording
'    End If
'
''Release ATE Safty net
'        With TheHdw.DCVI.Pins(sBuck1PFBPins)
'            .Gate = False
'            .Disconnect
'
'        End With
'
'        TheHdw.DCVI.Pins(sVddBuckPins).Voltage = 3.8
'        TheHdw.Utility.Pins("K1460,K1461,K1560,K1561,K1660,K1661").State = tlUtilBitOff
'
'If (bSupress) Then
'        g_RegVal = &H0: AHB_WRITE "TST_CTRL_DFT_FORCE", g_RegVal    'enter testmode
'        g_RegVal = &H0: AHB_WRITE "DVC_SCHEDULER_DVC_TEST_ALLOW_VSEL_WRITE_WHEN_DISABLED", g_RegVal
'
'End If
'----------------------------------Only Selected Site execute the default pattern---------------------------------------------------------
'Skip_DefPat:
'theexec.Sites.Selected = sbSaveSiteStatus
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'20200224 change the dsp function to a private local function
'___Replace elements in gD_wPGMData by the input value.
Private Function OTPData2DSPWave(ByVal m_lOTPIdx As Long)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "OTPData2DSPWave"
    Dim BinDspwave As New DSPWave
    Dim ValDSPWave As New DSPWave
    Dim m_vDecimal As Variant
    Dim m_lBitWidth As Long
    
    m_lBitWidth = g_OTPData.Category(m_lOTPIdx).lBitWidth
    m_vDecimal = g_OTPData.Category(m_lOTPIdx).Write.Value 'get the write value
    
    ValDSPWave.CreateConstant m_vDecimal, 1, DspLong
    BinDspwave = ValDSPWave.ConvertStreamTo(tldspSerial, m_lBitWidth, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
    gD_wPGMData.ReplaceElements g_OTPData.Category(m_lOTPIdx).wBitIndex, BinDspwave

Exit Function
ErrHandler:
    'Call TheExec.AddOutput("VBT_HDRobj encountered an error with STDTestOnProgStart.  More Info:" & vbCrLf & err.Description, vbBlue, False)
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function OTP_PostBurn_AHB_SetUp() As Long


End Function










