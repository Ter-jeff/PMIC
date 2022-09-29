Attribute VB_Name = "LIB_OTP_Common"
'T-AutoGen-Version : 1.3.0.1
'ProjectName_A1_TestPlan_20220226.xlsx
'ProjectName_A0_otp_AVA.otp
'ProjectName_A0_OTP_register_map.yaml
'ProjectName_A0_Pattern_List_Ext_20190823.csv
'ProjectName_A0_scgh_file#1_20200207.xlsx
'ProjectName_A0_VBTPOP_Gen_tool_MP10P_BuckSW_UVI80_DiffMeter_20200430.xlsm
Option Explicit

'*************************************************************
'           Common Use Functions
'*************************************************************


Public Function ConvertLotIdLetter2Bin(r_sInputStr As String) As String
    Dim sFuncName As String: sFuncName = "ConvertLotIdLetter2Bin"
    On Error GoTo ErrHandler
        
    Dim iStrIdx As Integer
    Dim iArrayIdx As Integer
    Dim sPerChar As String
    Dim sBinStr As String
    Dim sDecodeBin As String
    Dim asLotIDPrefix() As Variant
    Dim asBinary() As Variant

    r_sInputStr = UCase(r_sInputStr)

    asLotIDPrefix = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                      "A", "B", "C", "D", "E", "F", "G", "H", "I", _
                      "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                      "S", "T", "U", "V", "W", "X", "Y", "Z")

    asBinary = Array("000000", "000001", "000010", "000011", "000100", "000101", "000110", "000111", "001000", "001001", _
                    "001010", "001011", "001100", "001101", "001110", "001111", "010000", "010001", "010010", _
                    "010011", "010100", "010101", "010110", "010111", "011000", "011001", "011010", "011011", _
                    "011100", "011101", "011110", "011111", "100000", "100001", "100010", "100011")
                    
    sBinStr = ""
    For iStrIdx = 1 To Len(r_sInputStr)
        sPerChar = Mid(r_sInputStr, iStrIdx, 1)
        'One-to-One mapping, asLotIDPrefix() mappping to asBinary()
        For iArrayIdx = 0 To UBound(asLotIDPrefix)
            If (sPerChar = asLotIDPrefix(iArrayIdx)) Then
               sDecodeBin = asBinary(iArrayIdx)
               Exit For
            End If
        Next iArrayIdx
        sBinStr = sBinStr + sDecodeBin
    Next iStrIdx
    
    ''''Here sBinStr =[Ch1(MSB...LSB)][Ch2(MSB...LSB)][Ch3(MSB...LSB)][Ch4(MSB...LSB)][Ch5(MSB...LSB)][Ch6(MSB...LSB)]
    ConvertLotIdLetter2Bin = sBinStr

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ConvertLotIdBin2Letter(r_sBinStr As String) As String
    Dim sFuncName As String: sFuncName = "ConvertLotIdBin2Letter"
    On Error GoTo ErrHandler
    
    Dim lStrIdx As Long
    Dim asLotIDPrefix() As Variant
    Dim asBinary() As Variant
    
    '=== Initialization ===
    'for LotId
    asLotIDPrefix = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                   "A", "B", "C", "D", "E", "F", "G", "H", "I", _
                   "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                   "S", "T", "U", "V", "W", "X", "Y", "Z")
                   
    asBinary = Array("000000", "000001", "000010", "000011", "000100", "000101", "000110", "000111", "001000", "001001", _
                 "001010", "001011", "001100", "001101", "001110", "001111", "010000", "010001", "010010", _
                 "010011", "010100", "010101", "010110", "010111", "011000", "011001", "011010", "011011", _
                 "011100", "011101", "011110", "011111", "100000", "100001", "100010", "100011")
                          

    For lStrIdx = 0 To UBound(asLotIDPrefix)
      If r_sBinStr = asBinary(lStrIdx) Then
            ConvertLotIdBin2Letter = asLotIDPrefix(lStrIdx)
            Exit For
      End If
    Next lStrIdx

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ConvertFormat_Dec2Bin(ByVal lInputDec As Long, r_lBitWidth As Long, ByRef r_alBinary() As Long) As String
    Dim sFuncName As String: sFuncName = "ConvertFormat_Dec2Bin"
    On Error GoTo ErrHandler

    'Debug.Print "lInputDec = " & lInputDec & ", bitwidth=" & r_lBitWidth
    ''''-----------------------------------
    ''''<Example>
    ''''lInputDec = 11, bitwidth=6
    ''''r_alBinary [0] = 1
    ''''r_alBinary [1] = 1
    ''''r_alBinary [2] = 0
    ''''r_alBinary [3] = 1
    ''''r_alBinary [4] = 0
    ''''r_alBinary [5] = 0
    ''''bitstrM [MSB...LSB] = 001011
    ''''-----------------------------------

    Dim lBitIdx As Long
    Dim BitStrM As String
    
    ''Initialize the content of array
    ReDim r_alBinary(r_lBitWidth - 1) ''''r_alBinary[0] is LSB
    BitStrM = ""

    For lBitIdx = 0 To r_lBitWidth - 1
        r_alBinary(lBitIdx) = 0
        If (lInputDec Mod 2) Then
            r_alBinary(lBitIdx) = 1
        Else
            r_alBinary(lBitIdx) = 0
        End If
        BitStrM = CStr(r_alBinary(lBitIdx)) + BitStrM ''''[MSB...LSB]
        lInputDec = Fix(lInputDec / 2)
        ''Debug.Print "r_alBinary[" & lBitIdx & "] = " & r_alBinary(lBitIdx)
    Next lBitIdx
    ''Debug.Print "bitstrM[MSB...LSB] = " & bitstrM
    
    ConvertFormat_Dec2Bin = BitStrM

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ConvertFormat_Dec2Bin_Complement(ByVal vInputDec As Variant, ByVal lBitWidth As Long) As String
    Dim sFuncName As String: sFuncName = "ConvertFormat_Dec2Bin_Complement"
    On Error GoTo ErrHandler

    Dim lBitIdx As Long
    Dim sBitstrM As String: sBitstrM = ""
    Dim alBinary() As Long
    Dim sComplement As String
    ReDim alBinary(lBitWidth - 1) ''''alBinary[0] is LSB
    
    
    If (vInputDec > 2 ^ lBitWidth - 1) Then
       TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check the number and bit width."
       GoTo ErrHandler
    End If
    
    If (vInputDec >= 2 ^ 31) Then
        '2017/10/17
        If vInputDec > 2 ^ 32 - 1 Then
            GoTo ErrHandler
            TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check the number:" + CStr(vInputDec) + " > 2 ^ 32 - 1"
        End If
        sComplement = "1"
        vInputDec = vInputDec - 2 ^ 31
    End If
    
    For lBitIdx = 0 To lBitWidth - 1
        alBinary(lBitIdx) = 0
        If (vInputDec Mod 2) Then
            alBinary(lBitIdx) = 1
        Else
            alBinary(lBitIdx) = 0
        End If
        If sComplement = "1" And lBitIdx = 31 Then
            Exit For
        End If
        sBitstrM = CStr(alBinary(lBitIdx)) + sBitstrM ''''[MSB...LSB]
        vInputDec = Fix(vInputDec / 2)
    Next lBitIdx
    
    
    sBitstrM = sComplement + sBitstrM
    
    
    ConvertFormat_Dec2Bin_Complement = sBitstrM
    'Debug.Print sBitstrM
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


''''Here binstr (default) is [MSB......LSB]
''''update with the optional r_bBinStrMsb to True or False
Public Function ConvertFormat_Bin2Dec(r_sBinStr As String, Optional r_bBinStrMsb As Boolean = True) As Variant
    Dim sFuncName As String: sFuncName = "ConvertFormat_Bin2Dec"
    On Error GoTo ErrHandler
    
    ''''EX:
    ''''BinStr=11001, dec=25
    
    Dim lBitIdx As Long
    Dim lDec As Long: lDec = 0
    Dim dDbl As Double
    Dim lBitWidth As Long
    
    lBitWidth = Len(r_sBinStr)
    
    ''''case: r_sBinStr is [LSB...MSB]
    ''''Then set r_bBinStrMsb to Fasle, r_sBinStr should be reversed to [MSB...LSB]
    If (r_bBinStrMsb = False) Then
        r_sBinStr = StrReverse(r_sBinStr)
    End If

    ''''<NOTICE>
    ''''if lBitWidth >31, it will result in an overflow error message for (Clng).
    If (lBitWidth <= gD_slOTP_REGDATA_BW - 1) Then
        For lBitIdx = 0 To lBitWidth - 1
            lDec = lDec + CLng(Mid(r_sBinStr, lBitIdx + 1, 1)) * (2 ^ (lBitWidth - 1 - lBitIdx))
        Next lBitIdx
        ConvertFormat_Bin2Dec = lDec
    Else
        For lBitIdx = 0 To lBitWidth - 1
            dDbl = dDbl + CDbl(Mid(r_sBinStr, lBitIdx + 1, 1)) * CDbl(2 ^ (lBitWidth - 1 - lBitIdx))
        Next lBitIdx
        ConvertFormat_Bin2Dec = dDbl
    End If

    ''Debug.Print "BinStr=" + r_sBinStr + ", Dec=" + CStr(auto_OTP_BinStr2Dec)

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'CLAIRE 20170817
Public Function ConvertFormat_Bin2Hex(ByVal r_sBinStr As String, ByVal lHexBit As Long) As String
    Dim sFuncName As String: sFuncName = "ConvertFormat_Bin2Hex"
    On Error GoTo ErrHandler

    Dim lBitIdx As Integer
    Dim lBinStrLen As Long
    Dim iHexMOD As Integer
    Dim sHexStr As String
    Dim sHexVal As String
    Dim lHexStrLen As Long

    sHexStr = ""
    
    lBinStrLen = Len(r_sBinStr)
    If (lBinStrLen Mod (4)) > 0 Then
        lHexStrLen = (lBinStrLen \ 4) + 1
    Else
        lHexStrLen = lBinStrLen \ 4
    End If
    
    If lHexBit > lHexStrLen Then
        lHexStrLen = lHexBit
    End If

    iHexMOD = lHexStrLen * 4 - lBinStrLen
    
    If iHexMOD > 0 Then
        For lBitIdx = 0 To iHexMOD - 1
            r_sBinStr = "0" & r_sBinStr
        Next lBitIdx
    End If

    For lBitIdx = 0 To lHexStrLen - 1
        If Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "0000" Then
            sHexVal = "0"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "0001" Then
            sHexVal = "1"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "0010" Then
            sHexVal = "2"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "0011" Then
            sHexVal = "3"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "0100" Then
            sHexVal = "4"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "0101" Then
            sHexVal = "5"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "0110" Then
            sHexVal = "6"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "0111" Then
            sHexVal = "7"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "1000" Then
            sHexVal = "8"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "1001" Then
            sHexVal = "9"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "1010" Then
            sHexVal = "A"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "1011" Then
            sHexVal = "B"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "1100" Then
            sHexVal = "C"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "1101" Then
            sHexVal = "D"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "1110" Then
            sHexVal = "E"
        ElseIf Mid(r_sBinStr, lBitIdx * 4 + 1, 4) = "1111" Then
            sHexVal = "F"
        Else
            sHexVal = "X"
        End If

        sHexStr = sHexStr & sHexVal
    Next lBitIdx

    ConvertFormat_Bin2Hex = sHexStr
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
' Gen tool need this function "ToSiteLong"
Public Function ToSiteLong(ByVal vValue As Variant) As SiteLong
    Dim sFuncName As String: sFuncName = "ToSiteLong"
    On Error GoTo ErrHandler
    Dim slValue As New SiteLong
    Dim lInValuefmt As Long
    
    If (VarType(vValue) = vbLong) Or (VarType(vValue) = vbInteger) Then
        slValue = vValue
    ElseIf (VarType(vValue) = vbString) Then
        lInValuefmt = InStr(1, vValue, "0x")
        If (lInValuefmt > 0) Then
            slValue = CLng("&H" & Mid(vValue, lInValuefmt + 2))
        ElseIf (InStr(1, vValue, "&H") = 0) Then
            slValue = CLng("&H" & vValue)
        Else
            slValue = CLng(vValue)
        End If
    Else
        slValue = vValue
    End If
    
    Set ToSiteLong = slValue
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ConvertECID2Bin(r_vLotID As Variant, r_lWfID As Long, r_lXcoor As Long, r_lYcoor As Long, Optional r_lRev As Long = 0) As String
    Dim sFuncName As String: sFuncName = "ConvertECID2Bin"
    On Error GoTo ErrHandler
    Dim sBinStr As String
    Dim alBinary() As Long
    
    '< Step1. transfer lotid, waferid, x and y coord to binary string>
    sBinStr = ""
    sBinStr = sBinStr + ConvertLotIdLetter2Bin(CStr(r_vLotID)) ''''LotID  36 bits (6x6)
    sBinStr = sBinStr + ConvertFormat_Dec2Bin(r_lWfID, 5, alBinary)  ''''WaferID 5 bits
    sBinStr = sBinStr + ConvertFormat_Dec2Bin(r_lXcoor, 6, alBinary)  ''''X Coord 6 bits
    sBinStr = sBinStr + ConvertFormat_Dec2Bin(r_lYcoor, 6, alBinary) ''''Y Coord 6 bits
    
    Do Until (Len(sBinStr) Mod 56 = 0) ''''8bits X 7 Regs = 56
        sBinStr = sBinStr + "0"
    Loop
    
    ''''OTP_Revision[58:56]
    sBinStr = sBinStr + "000"     'auto_OTP_Dec2Bin(r_lRev, 3, alBinary)  ''''OTP_Revision 3 bits '2018/08/29 Remove gC_OTP_Revision
    sBinStr = sBinStr + "00000"   '2018/08/29 TTR
    Do Until (Len(sBinStr) Mod 64 = 0) ''''8bits X 8 Regs = 64
        sBinStr = sBinStr + "0"
    Loop
    
    ''''Here total 64bits
    ''''[LotID_36bits(MSB...LSB)][WaferID_5bits(MSB...LSB)][XCoord_6bits(MSB...LSB)][YCoord_6bits(MSB...LSB)][000][Rev_3bits(MSB...LSB)][00000]
    ConvertECID2Bin = sBinStr

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function FormatLog(r_vNum As Variant, r_ilength As Integer) As String
    Dim sFuncName As String: sFuncName = "FormatLog"
    On Error GoTo ErrHandler
    
    ''''Example
    ''''----------------------------------------
    '''' r_ilength > 0  is to right shift
    '''' r_ilength < 0  is to left  shift
    ''''----------------------------------------
    ''''FormatLog(123456, 8) + "...end"
    ''''  123456...end
    ''''
    ''''FormatLog(123456,-8) + "...end"
    ''''123456  ...end
    ''''
    ''''----------------------------------------
    
    Dim sNum As String
    Dim lNumLen As Long
    Dim lSpcLen As Long
    
    sNum = CStr(r_vNum)
    lNumLen = Len(sNum)
    
    If (lNumLen > Abs(r_ilength)) Then
        lSpcLen = 0
    Else
        lSpcLen = Abs(r_ilength) - lNumLen
    End If
    
    If (r_ilength < 0) Then   ''''number shift to the very left
        FormatLog = sNum + Space(lSpcLen)
    Else ''''default: shift to the very right
        FormatLog = Space(lSpcLen) + sNum
    End If

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ResetDspGlobalVariable() As Long
    Dim sFuncName As String: sFuncName = "ResetDspGlobalVariable"
    On Error GoTo ErrHandler

    ''''---------------------------------------------
    '''' Clear the global DSP DSPWave
    ''''---------------------------------------------
    If TheExec.Sites.Active.Count = 0 Then Exit Function 'to prevent to clear again during job running 20190522
    
    ''''20200313 this function is called in OnDSPGlobalVariableReset(), don't need the site-loop
    For Each Site In TheExec.Sites
        gD_wPGMData.Clear
        gD_wReadData.Clear
        'gDW_DefaultDSPRawData.Clear
        gD_wDEIDPGMBits.Clear
        gD_wCRCSelfLUT.Clear
        'DefaultReal
        gDW_RealDef_fromWrite.Clear
    Next Site
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function IsSheetExists(r_sSheetName As String) As Boolean
    Dim sFuncName As String: sFuncName = "isSheetExists"
    On Error GoTo ErrHandler
    
    IsSheetExists = (Sheets(r_sSheetName).Name <> "")
        
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'20190526
'
'Public Function STT()
''    If g_bDetailedTestTimeProfile = True Then
'        TheHdw.StartStopwatch
''    End If
'End Function
Public Function OTP_SPT_D(Optional r_sMessage As String = "Exe Time = ", Optional r_bEnforceLogTime As Boolean = True)
    Dim sFuncName As String: sFuncName = "OTP_SPT_D"
    On Error GoTo ErrHandler

    Dim dExetime As Double
    If g_bTestTimeProfileDebugPrint = True Or r_bEnforceLogTime = True Then
        dExetime = TheHdw.ReadStopwatch
        TheExec.Datalog.WriteComment r_sMessage & dExetime
    End If
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetPatListFromPatternSet_OTP(r_sTestPat As String, r_asRtnPatNames() As String, Optional r_lRtnPatCnt As Long) As String
    Dim sFuncName As String: sFuncName = "GetPatListFromPatternSet_OTP"
    On Error GoTo ErrHandler
                             
    Dim lPatCnt As Long                          '<- Number of patterns in set
    Dim asRtnPatNames1() As String
    Dim asRtnPatNames2() As String
    Dim lIdx As Long, lIdx2 As Long
    
    '___ Init _____________________________________________________________________________
'    On Error GoTo errhandler
    
    '___ Check the name ___________________________________________________________________
    '    Individual pattern name or non-pattern string returns an error - thus false
    '--------------------------------------------------------------------------------------
    r_asRtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(r_sTestPat, lPatCnt)
    If lPatCnt = 0 Then ''''20200313, need to confirm here
        r_lRtnPatCnt = 0
        Exit Function
    End If
    If (UBound(r_asRtnPatNames) > 0) Then
        If LCase(r_asRtnPatNames(0)) Like "*.pat*" Then
            'PATT_GetPatListFromPatternSet = True
            r_lRtnPatCnt = UBound(r_asRtnPatNames) + 1
        Else
            r_lRtnPatCnt = 0
            For lIdx = 0 To UBound(r_asRtnPatNames)
                asRtnPatNames2 = TheExec.DataManager.Raw.GetPatternsInSet(r_asRtnPatNames(lIdx), lPatCnt)
                r_lRtnPatCnt = r_lRtnPatCnt + UBound(asRtnPatNames2) + 1
            Next lIdx
            asRtnPatNames1 = TheExec.DataManager.Raw.GetPatternsInSet(r_sTestPat, lPatCnt)
            ReDim r_asRtnPatNames(r_lRtnPatCnt - 1)
            r_lRtnPatCnt = 0
            For lIdx = 0 To UBound(asRtnPatNames1)
                asRtnPatNames2 = TheExec.DataManager.Raw.GetPatternsInSet(asRtnPatNames1(lIdx), lPatCnt)
                For lIdx2 = 0 To UBound(asRtnPatNames2)
                    If LCase(asRtnPatNames2(lIdx2)) Like "*.pat*" Then
                        r_asRtnPatNames(r_lRtnPatCnt) = asRtnPatNames2(lIdx2)
                    Else
                        TheExec.ErrorLogMessage r_sTestPat & " in more than 2 level of pattern set"
                    End If
                    r_lRtnPatCnt = r_lRtnPatCnt + 1
                Next lIdx2
            Next lIdx
            'PATT_GetPatListFromPatternSet = True
        End If
    Else
        If LCase(r_asRtnPatNames(0)) Like "*.pat*" Then
            'PATT_GetPatListFromPatternSet = True
            r_lRtnPatCnt = 1
        Else
            r_asRtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(r_asRtnPatNames(0), lPatCnt)
            r_lRtnPatCnt = UBound(r_asRtnPatNames) + 1
            For lIdx2 = 0 To UBound(r_asRtnPatNames)
                If LCase(r_asRtnPatNames(lIdx2)) Like "*.pat*" Then
                Else
                    TheExec.ErrorLogMessage r_sTestPat & " in more than 2 level of pattern set"
                End If
            Next lIdx2
        End If
    End If
    
          GetPatListFromPatternSet_OTP = r_asRtnPatNames(0)
          
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'Public Function OTP_Execute_Pattern(PatternName As PatternSet, Optional mPrintLog As Boolean = True)
'On Error GoTo ErrHandler
'    Dim sFuncName As String: sFuncName = "OTP_Execute_Pattern"
'
'    TheHdw.Patterns(PatternName).Load
'    TheHdw.Wait 0.001
'    If mPrintLog Then TheExec.Datalog.WriteComment ("RUN PAT:" & PatternName.Value)
'    TheHdw.Patterns(PatternName).Start
'    TheHdw.Digital.Patgen.HaltWait
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'20190329
Public Function GetPatListNExecutePat(r_psTestPat As PatternSet)
    Dim sFuncName As String: sFuncName = "GetPatListNExecutePat"
    On Error GoTo ErrHandler
    Dim asPatt() As String
    Dim lPatCnt As Long
    
     r_psTestPat.Value = GetPatListFromPatternSet_OTP(r_psTestPat.Value, asPatt, lPatCnt)
     TheHdw.Patterns(r_psTestPat.Value).Load
     TheHdw.Patterns(r_psTestPat.Value).Start
     TheHdw.Digital.Patgen.HaltWait
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function MeasureVbyPPMU(ByVal sPinName As String) As SiteDouble
    Dim sFuncName As String: sFuncName = "MeasureVbyPPMU"
    On Error GoTo ErrHandler
    Dim sdMeasVal As New SiteDouble
    TheHdw.Digital.Pins(sPinName).Disconnect

    With TheHdw.PPMU.Pins(sPinName)
        .Gate = tlOff
        .ForceI 0#
        .ClampVHi = 6
        .ClampVLo = 0
        Call .Connect
        .Gate = True
    End With
    Call TheHdw.Wait(0.001)
    sdMeasVal = TheHdw.PPMU.Pins(sPinName).Read(tlPPMUReadMeasurements, 10)
     
     Set MeasureVbyPPMU = sdMeasVal
    'Reset:
    TheHdw.PPMU.Pins(sPinName).Disconnect
    TheHdw.Digital.Pins(sPinName).Connect
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub ForceVbyPPMU(ByVal sPinName As String, ByVal dForceLevel As Double)
    Dim sFuncName As String: sFuncName = "ForceVbyPPMU"
    On Error GoTo ErrHandler
    TheHdw.Digital.Pins(sPinName).Disconnect

    With TheHdw.PPMU.Pins(sPinName)
        .Gate = tlOff
        Call .ForceV(dForceLevel, 0.002)
        .ClampVHi = 6
        .ClampVLo = 0
        Call .Connect
        .Gate = tlOn
    End With
   
   TheHdw.Wait 1 * ms
   
Exit Sub
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Public Sub MakeFolder(r_sFolderName As String)
    Dim sFuncName As String: sFuncName = "r_sFolderName"
    On Error GoTo ErrHandler
    Dim oFile As Object, ofso As Object
    Dim vCurrFolder As Variant
    vCurrFolder = CurDir() & r_sFolderName
    'vCurrFolder = Application.ActiveWorkbook.Path & "\OTPData"
    
    Set ofso = CreateObject("Scripting.FileSystemObject")
    If Not ofso.FolderExists(vCurrFolder) Then Set oFile = ofso.CreateFolder(vCurrFolder)
    
Exit Sub
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'Sub RemoveWraps()
'On Error GoTo ErrHandler
'Dim sFuncName As String: sFuncName = "RemoveWraps"
''the wrap text makes the line messed up
'Dim Rng As Range
'Dim WorkRng As Range
'Dim worksheet_name As String
'Set WorkRng = Application.Selection
'worksheet_name = "AHB_register_map"
'Set WorkRng = Application.InputBox("Range", worksheet_name, WorkRng.Address, Type:=8)
'For Each Rng In WorkRng
'If Rng.Value Like "*f(x)*" Then
'Rng.Value = Replace(Rng.Value, Chr(10), "")
'End If
'Next
'
'Exit Sub
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Sub Else Resume Next
'End Sub
Public Function HEADERINFO()
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "HEADERINFO"
Dim strFlowCtr As String
Dim strEnableWrds As String
strEnableWrds = ""
Dim strNotEnableWrds As String
strNotEnableWrds = ""

Dim strInSwVer As String
Dim strInSwBuild As String
Dim FlowCtr As Long
Dim SoftwareVersion As Long
Dim IGXLbuild As Long


TheExec.Datalog.WriteComment "COMPUTERDATA"
TheExec.Datalog.WriteComment "-----------------------------------------------"
TheExec.Datalog.WriteComment "ComputerName: " & TheHdw.Computer.Name  '
TheExec.Datalog.WriteComment "OS: " & TheHdw.Computer.OperatingSystem  ' service pack
TheExec.Datalog.WriteComment "CPUnumbers: " & TheHdw.Computer.NumberofProcessors  '
TheExec.Datalog.WriteComment "Memory: " & TheHdw.Computer.PhysicalMemory  '
TheExec.Datalog.WriteComment "IS3GENABLED: " & TheHdw.Computer.Is3GEnabled  '
TheExec.Datalog.WriteComment "CPUSPEED: " & TheHdw.Computer.ProcessorSpeed  '
TheExec.Datalog.WriteComment "TYPEOFCPU: " & TheHdw.Computer.ProcessorType  '


TheExec.Datalog.WriteComment vbNullString

TheExec.Datalog.WriteComment ""
TheExec.Datalog.WriteComment "***************************************************"
DSPinfo
TheExec.Datalog.WriteComment vbNullString


TheExec.Datalog.WriteComment "User DATA"
TheExec.Datalog.WriteComment "***************************************************"
TheExec.Datalog.WriteComment "User: " & TheHdw.Computer.UserName  '
TheExec.Datalog.WriteComment vbNullString

TheExec.Datalog.WriteComment "IGXL DATA"
TheExec.Datalog.WriteComment "***************************************************"
TheExec.Datalog.WriteComment "IGXLVersion: " & TheExec.SoftwareVersion  '
TheExec.Datalog.WriteComment "IGXLBuild: " & TheExec.SoftwareBuild  '
TheExec.Datalog.WriteComment vbNullString


' Generate a string based on current program "flow-control" settings...i.e. job, channelmap, part, env...enable words
TheExec.Datalog.WriteComment "Flow SELECTION DATA"
TheExec.Datalog.WriteComment "***************************************************"
TheExec.Datalog.WriteComment ("CurrentJob: " & TheExec.CurrentJob)
TheExec.Datalog.WriteComment ("CurrentChanMap: " & TheExec.CurrentChanMap)
TheExec.Datalog.WriteComment ("CurrentPart: " & TheExec.CurrentPart)
TheExec.Datalog.WriteComment ("CurrentEnv: " & TheExec.CurrentEnv)
'TheExec.Datalog.WriteComment ("NameofWorkbook: " & ActiveWorkbook.Name)
TheExec.Datalog.WriteComment ("NameofWorkbook: " & TheExec.TestProgram.Name) '20180920 Modified to get workbook name in IGXL 9.0
'TheExec.Datalog.WriteComment ("Revisionoftheprogram: " & TheExec.Datalog.Setup.LotSetup.JobRev)


' Get enable word status...
FlowCtr = getflowid(strEnableWrds, strNotEnableWrds)

TheExec.Datalog.WriteComment ("EnableWordsSET: " & strEnableWrds)
TheExec.Datalog.WriteComment ("EnableWordsNOTSET: " & strNotEnableWrds)
TheExec.Datalog.WriteComment vbNullString

' Get software version...
strInSwVer = TheExec.SoftwareVersion
strInSwBuild = TheExec.SoftwareBuild


'Call TheExec.Datalog.WriteComment("ProgramFlowControl: " & strIn)
Call TheExec.Datalog.WriteComment("IGXLSWVersion: " & strInSwVer)
Call TheExec.Datalog.WriteComment("IGXLSWBuild: " & strInSwBuild)


Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function




