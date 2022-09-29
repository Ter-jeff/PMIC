Attribute VB_Name = "DATALOG_HEADER"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit


Public HDRStdDatalogSetupComplete As Boolean
Private HDRStdFlowIDHashCurrent As Long
Private Const MD5_BLK_LEN As Long = 64
Private Const S11 As Long = 7
Private Const S12 As Long = 12
Private Const S13 As Long = 17
Private Const S14 As Long = 22
Private Const S21 As Long = 5
Private Const S22 As Long = 9
Private Const S23 As Long = 14
Private Const S24 As Long = 20
Private Const S31 As Long = 4
Private Const S32 As Long = 11
Private Const S33 As Long = 16
Private Const S34 As Long = 23
Private Const S41 As Long = 6
Private Const S42 As Long = 10
Private Const S43 As Long = 15
Private Const S44 As Long = 21
' Constants for unsigned word addition
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647



Public InitWordsEnabled As String




'Public Function STDTestOnProgStart()
'
'    On Error GoTo ErrorHandler
'
'    Call HEADERINFO    ' Get the flow control words, calc a hash and record it to datalog/stdf
'
'    Exit Function
'ErrorHandler:
'    Call TheExec.AddOutput("VBT_HDRobj encountered an error with STDTestOnProgStart.  More Info:" & vbCrLf & err.Description, vbBlue, False)
'
'
'End Function




Public Sub DSPinfo()

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "DSPinfo"


    Dim idx       As Long
    Dim Count     As Long
    Dim strTmp    As String

    strTmp = "thehdw.dsp.computers("
    Count = TheHdw.DSP.Computers.Count
    TheExec.Datalog.WriteComment "the Tester has " & Count & " DSP machines"

    For idx = 1 To Count
        TheExec.Datalog.WriteComment strTmp & idx & ").NumberOfCoresPerCPU = " & TheHdw.DSP.Computers(idx).NumberOfCoresPerProcessor
        TheExec.Datalog.WriteComment strTmp & idx & ").OperatingSystem = " & TheHdw.DSP.Computers(idx).OperatingSystem
        TheExec.Datalog.WriteComment strTmp & idx & ").PhysicalMemory = " & TheHdw.DSP.Computers(idx).PhysicalMemory
        TheExec.Datalog.WriteComment strTmp & idx & ").CPUSpeed = " & TheHdw.DSP.Computers(idx).ProcessorSpeed
        TheExec.Datalog.WriteComment strTmp & idx & ").CPUType = " & TheHdw.DSP.Computers(idx).ProcessorType
        TheExec.Datalog.WriteComment strTmp & idx & ").IPAddresss = " & TheHdw.DSP.Computers(idx).IPAddress

    Next idx

    Exit Sub
ErrHandler:
    RunTimeError funcName
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Sub Else Resume Next

End Sub

'Public Function DatalogModuleInfo()
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "DatalogModuleInfo"
'
'
'    ' Datalog
'    TheExec.Datalog.WriteComment "Module Information"
'    TheExec.Datalog.WriteComment "-----------------------------------------------"
'    'TheExec.Datalog.WriteComment CStr(LIBNAME) & "_Rev: " & CStr(SOURCEREV)  '
'    'TheExec.Datalog.WriteComment CStr(LIBNAME) & "_Date: " & SOURCEDATE
'    TheExec.Datalog.WriteComment vbNullString
'
'    Exit Function
'
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function



'Public Function HEADERINFO()
'
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "HEADERINFO"
'
'
'    Dim strFlowCtr As String
'    Dim strEnableWrds As String
'    strEnableWrds = ""
'    Dim strNotEnableWrds As String
'    strNotEnableWrds = ""
'
'    Dim strInSwVer As String
'    Dim strInSwBuild As String
'    Dim FlowCtr   As Long
'    Dim SoftwareVersion As Long
'    Dim IGXLbuild As Long
'
'
'    TheExec.Datalog.WriteComment "COMPUTERDATA"
'    TheExec.Datalog.WriteComment "-----------------------------------------------"
'    TheExec.Datalog.WriteComment "ComputerName: " & TheHdw.Computer.Name  '
'    TheExec.Datalog.WriteComment "OS: " & TheHdw.Computer.OperatingSystem  ' service pack
'    TheExec.Datalog.WriteComment "CPUnumbers: " & TheHdw.Computer.NumberofProcessors  '
'    TheExec.Datalog.WriteComment "Memory: " & TheHdw.Computer.PhysicalMemory  '
'    TheExec.Datalog.WriteComment "IS3GENABLED: " & TheHdw.Computer.Is3GEnabled  '
'    TheExec.Datalog.WriteComment "CPUSPEED: " & TheHdw.Computer.ProcessorSpeed  '
'    TheExec.Datalog.WriteComment "TYPEOFCPU: " & TheHdw.Computer.ProcessorType  '
'
'
'    TheExec.Datalog.WriteComment vbNullString
'
'    TheExec.Datalog.WriteComment ""
'    TheExec.Datalog.WriteComment "***************************************************"
'    DSPinfo
'    TheExec.Datalog.WriteComment vbNullString
'
'
'    TheExec.Datalog.WriteComment "User DATA"
'    TheExec.Datalog.WriteComment "***************************************************"
'    TheExec.Datalog.WriteComment "User: " & TheHdw.Computer.UserName  '
'    TheExec.Datalog.WriteComment vbNullString
'
'    TheExec.Datalog.WriteComment "IGXL DATA"
'    TheExec.Datalog.WriteComment "***************************************************"
'    TheExec.Datalog.WriteComment "IGXLVersion: " & TheExec.SoftwareVersion  '
'    TheExec.Datalog.WriteComment "IGXLBuild: " & TheExec.SoftwareBuild  '
'    TheExec.Datalog.WriteComment vbNullString
'
'
'    ' Generate a string based on current program "flow-control" settings...i.e. job, channelmap, part, env...enable words
'    TheExec.Datalog.WriteComment "Flow SELECTION DATA"
'    TheExec.Datalog.WriteComment "***************************************************"
'    TheExec.Datalog.WriteComment ("CurrentJob: " & TheExec.CurrentJob)
'    TheExec.Datalog.WriteComment ("CurrentChanMap: " & TheExec.CurrentChanMap)
'    TheExec.Datalog.WriteComment ("CurrentPart: " & TheExec.CurrentPart)
'    TheExec.Datalog.WriteComment ("CurrentEnv: " & TheExec.CurrentEnv)
'    TheExec.Datalog.WriteComment ("NameofWorkbook: " & TheExec.TestProgram.Name)
'    'TheExec.Datalog.WriteComment ("Revisionoftheprogram: " & TheExec.Datalog.Setup.LotSetup.JobRev)
'
'
'    ' Get enable word status...
'    FlowCtr = getflowid(strEnableWrds, strNotEnableWrds)
'
'    TheExec.Datalog.WriteComment ("EnableWordsSET: " & strEnableWrds)
'    TheExec.Datalog.WriteComment ("EnableWordsNOTSET: " & strNotEnableWrds)
'    TheExec.Datalog.WriteComment vbNullString
'
'    ' Get software version...
'    strInSwVer = TheExec.SoftwareVersion
'    strInSwBuild = TheExec.SoftwareBuild
'
'
'    'Call TheExec.Datalog.WriteComment("ProgramFlowControl: " & strIn)
'    Call TheExec.Datalog.WriteComment("IGXLSWVersion: " & strInSwVer)
'    Call TheExec.Datalog.WriteComment("IGXLSWBuild: " & strInSwBuild)
'
'
'    Call Read_TesterConfig
'
'    Exit Function
'
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function

Public Function getflowid(ByRef strEnableWrds As String, ByRef strNotEnableWrds As String) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "getflowid"

    Dim strIn     As String
    Dim words()   As String
    Dim lc_words() As String
    Dim num_words As Long

    num_words = tl_ExecGetEnableWords(words)
    If num_words > 0 Then  ' do I have at least one enable word?

        ' set everything to lower case first...
        ReDim lc_words(num_words - 1)

        ' set enable words to lowercase and sort...
        Dim word  As Long
        For word = LBound(words) To UBound(words)
            lc_words(word) = LCase(words(word))
        Next
        Call QuickSort1(lc_words)  ' not sure that this is necessary...list seems pre-sorted

        For word = LBound(lc_words) To UBound(lc_words)
            If (TheExec.EnableWord(lc_words(word))) Then
                strEnableWrds = strEnableWrds & lc_words(word) & "|"
            Else
                strNotEnableWrds = strNotEnableWrds & lc_words(word) & "|"
            End If
        Next
    End If

    strIn = LCase(TheExec.CurrentJob) & "|" & LCase(TheExec.CurrentChanMap) & "|" & LCase(TheExec.CurrentPart) _
            & "|" & LCase(TheExec.CurrentEnv) & "|" & LCase(ActiveWorkbook.Name)

    getflowid = Val("&h" & Right$(MD5_string(strEnableWrds & strIn), 6))


    Exit Function

ErrHandler:
    'Debug.Print "getflowid  : " & err.Number & " : " & err.Description
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



'Public Sub HDRStdDatalogSetup()
'    On Error GoTo ErrorHandler
'
'    Dim strEnableWrds As String
'    Dim strNotEnableWrds As String
'
'    Dim HDRStdFlowIDHashNew As Long
'    HDRStdFlowIDHashNew = getflowid(strEnableWrds, strNotEnableWrds)
'
'    InitWordsEnabled = strEnableWrds
'
'    If (HDRStdFlowIDHashCurrent <> HDRStdFlowIDHashNew) Or TheExec.Datalog.Setup.LotSetup.DeviceNumber = 1 Then
'
'        HDRStdFlowIDHashCurrent = HDRStdFlowIDHashNew
'
'
'        ' Clean up and log Test Program Revision (Format: MMDDYY) and store in MIR
'        TheExec.Datalog.Setup.LotSetup.bJobRev = True
'
'        ' Add test condition/Insertion into MIR
'        TheExec.Datalog.Setup.LotSetup.bTestCode = True
'
'        TheExec.Datalog.Setup.LotSetup.ExecType = "IGXL"
'        TheExec.Datalog.Setup.LotSetup.bExecType = True
'
'        TheExec.Datalog.Setup.LotSetup.ExecVer = TheExec.SoftwareVersion & " (build: " & TheExec.SoftwareBuild & ")"
'        TheExec.Datalog.Setup.LotSetup.bExecVer = True
'
'        ' conditional statement added for HDR89335 PAT.
'        If (TheExec.CurrentPart <> "") Then
'            TheExec.Datalog.Setup.LotSetup.PartType = TheExec.CurrentPart
'        Else
'            TheExec.Datalog.Setup.LotSetup.PartType = TheExec.CurrentJob
'        End If
'
'
'
'        ' Set MIR::SUPERVISOR to test so that Optimal Test can easily idendify test data.
'        TheExec.Datalog.Setup.LotSetup.Supervisor = "test"
'        TheExec.Datalog.Setup.LotSetup.bSupervisor = True
'
'        TheExec.Datalog.Setup.SummarySetup.ReportPatternName = True      'Supported starting in IGXL 6.00.07
'
'        'The following DatalogSetup options are for generating Image compatible STDF
'        TheExec.Datalog.Setup.DatalogSetup.ReversePTRTestTxt = True
'        TheExec.Datalog.Setup.DatalogSetup.DisableChannelNumberInPTR = True
'        TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
'        TheExec.Datalog.Setup.DatalogSetup.TSRTestNamSameAsPTRTestTxt = True
'
'
'        ' Custom width of the datalog so its readable
'        ''''        theexec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
'        ''''        theexec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testname.Enable = True
'        ''''        theexec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testname.Width = 80
'        ''''        theexec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 80
'
'        'Turn Off Default PinlistData Datalogging Sequence
'        ''Call tl_PinListDataSort(False)
'
'        TheExec.Datalog.ApplySetup
'
'        HDRStdDatalogSetupComplete = True
'
'    End If
'
'
'
'    Exit Sub
'ErrorHandler:
'    Call TheExec.AddOutput("VBT_HDRobj encountered an error with HDRStdDatalogSetup.  More Info:" & vbCrLf & err.Description, vbBlue, False)
'
'End Sub

'Private Sub PrintDatalogHeader()
'
'
'    With TheExec.Datalog
'        .WriteComment ("**                                                                          **")
'        .WriteComment ("**                          TEST PROGRAM HEADEAR                            **")
'
'    End With
'
'End Sub


' Omit plngLeft & plngRight; they are used internally during recursion
Public Sub QuickSort1(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "QuickSort1"


    Dim lngFirst  As Long
    Dim lngLast   As Long
    Dim varMid    As Variant
    Dim varSwap   As Variant

    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        Do While pvarArray(lngFirst) < varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid < pvarArray(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then QuickSort1 pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then QuickSort1 pvarArray, lngFirst, plngRight

    Exit Sub
ErrHandler:
    RunTimeError funcName
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Sub Else Resume Next

End Sub


Public Function MD5_string(strMessage As String) As String
    ' Returns 32-char hex string representation of message digest
    ' Input as a string (max length 2^29-1 bytes)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "MD5_string"


    Dim abMessage() As Byte
    Dim mLen      As Long
    ' Cope with the empty string
    If Len(strMessage) > 0 Then
        abMessage = StrConv(strMessage, vbFromUnicode)
        ' Compute length of message in bytes
        mLen = UBound(abMessage) - LBound(abMessage) + 1
    End If
    MD5_string = MD5_bytes(abMessage, mLen)

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function MD5_bytes(abMessage() As Byte, mLen As Long) As String
    ' Returns 32-char hex string representation of message digest
    ' Input as an array of bytes of length mLen bytes
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "MD5_bytes"



    Dim nBlks     As Long
    Dim nBits     As Long
    Dim Block(MD5_BLK_LEN - 1) As Byte
    Dim State(3)  As Long
    Dim wb(3)     As Byte
    Dim sHex      As String
    Dim Index     As Long
    Dim partLen   As Long
    Dim i         As Long
    Dim j         As Long

    ' Catch length too big for VB arithmetic (268 million!)
    If mLen >= &HFFFFFFF Then Error 6     ' overflow

    ' Initialise
    ' Number of complete 512-bit/64-byte blocks to process
    nBlks = mLen \ MD5_BLK_LEN

    ' Load magic initialization constants
    State(0) = &H67452301
    State(1) = &HEFCDAB89
    State(2) = &H98BADCFE
    State(3) = &H10325476

    ' Main loop for each complete input block of 64 bytes
    Index = 0
    For i = 0 To nBlks - 1
        Call md5_transform(State, abMessage, Index)
        Index = Index + MD5_BLK_LEN
    Next

    ' Construct final block(s) with padding
    partLen = mLen Mod MD5_BLK_LEN
    Index = nBlks * MD5_BLK_LEN
    For i = 0 To partLen - 1
        Block(i) = abMessage(Index + i)
    Next
    Block(partLen) = &H80
    ' Make sure padding (and bit-length) set to zero
    For i = partLen + 1 To MD5_BLK_LEN - 1
        Block(i) = 0
    Next
    ' Two cases: partLen is < or >= 56
    If partLen >= MD5_BLK_LEN - 8 Then
        ' Need two blocks
        Call md5_transform(State, Block, 0)
        For i = 0 To MD5_BLK_LEN - 1
            Block(i) = 0
        Next
    End If
    ' Append number of bits in little-endian order
    nBits = mLen * 8
    Block(MD5_BLK_LEN - 8) = nBits And &HFF
    Block(MD5_BLK_LEN - 7) = nBits \ &H100 And &HFF
    Block(MD5_BLK_LEN - 6) = nBits \ &H10000 And &HFF
    Block(MD5_BLK_LEN - 5) = nBits \ &H1000000 And &HFF
    ' (NB we don't try to cope with number greater than 2^31)

    ' Final padded block with bit length
    Call md5_transform(State, Block, 0)

    ' Decode 4 x 32-bit words into 16 bytes with LSB first each time
    ' and return result as a hex string
    MD5_bytes = ""
    For i = 0 To 3
        Call uwSplit(State(i), wb(3), wb(2), wb(1), wb(0))
        For j = 0 To 3
            If wb(j) < 16 Then
                sHex = "0" & Hex(wb(j))
            Else
                sHex = Hex(wb(j))
            End If
            MD5_bytes = MD5_bytes & sHex
        Next
    Next

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


Private Sub md5_transform(State() As Long, buf() As Byte, ByVal Index As Long)
    ' Updates 4 x 32-bit values in state
    ' Input: the next 64 bytes in buf starting at offset index
    ' Assumes at least 64 bytes are present after offset index
    
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "md5_transform"

    Dim A         As Long
    Dim b         As Long
    Dim C         As Long
    Dim D         As Long
    Dim j         As Integer
    Dim x(15)     As Long
    

    A = State(0)
    b = State(1)
    C = State(2)
    D = State(3)

    ' Decode the next 64 bytes into 16 words with LSB first
    For j = 0 To 15
        x(j) = uwJoin(buf(Index + 3), buf(Index + 2), buf(Index + 1), buf(Index))
        Index = Index + 4
    Next

    ' Round 1
    A = ff(A, b, C, D, x(0), S11, &HD76AA478)   ' 1
    D = ff(D, A, b, C, x(1), S12, &HE8C7B756)   ' 2
    C = ff(C, D, A, b, x(2), S13, &H242070DB)   ' 3
    b = ff(b, C, D, A, x(3), S14, &HC1BDCEEE)   ' 4
    A = ff(A, b, C, D, x(4), S11, &HF57C0FAF)   ' 5
    D = ff(D, A, b, C, x(5), S12, &H4787C62A)   ' 6
    C = ff(C, D, A, b, x(6), S13, &HA8304613)   ' 7
    b = ff(b, C, D, A, x(7), S14, &HFD469501)   ' 8
    A = ff(A, b, C, D, x(8), S11, &H698098D8)   ' 9
    D = ff(D, A, b, C, x(9), S12, &H8B44F7AF)   ' 10
    C = ff(C, D, A, b, x(10), S13, &HFFFF5BB1)  ' 11
    b = ff(b, C, D, A, x(11), S14, &H895CD7BE)  ' 12
    A = ff(A, b, C, D, x(12), S11, &H6B901122)  ' 13
    D = ff(D, A, b, C, x(13), S12, &HFD987193)  ' 14
    C = ff(C, D, A, b, x(14), S13, &HA679438E)  ' 15
    b = ff(b, C, D, A, x(15), S14, &H49B40821)  ' 16

    ' Round 2
    A = GG(A, b, C, D, x(1), S21, &HF61E2562)   ' 17
    D = GG(D, A, b, C, x(6), S22, &HC040B340)   ' 18
    C = GG(C, D, A, b, x(11), S23, &H265E5A51)  ' 19
    b = GG(b, C, D, A, x(0), S24, &HE9B6C7AA)   ' 20
    A = GG(A, b, C, D, x(5), S21, &HD62F105D)   ' 21
    D = GG(D, A, b, C, x(10), S22, &H2441453)   ' 22
    C = GG(C, D, A, b, x(15), S23, &HD8A1E681)  ' 23
    b = GG(b, C, D, A, x(4), S24, &HE7D3FBC8)   ' 24
    A = GG(A, b, C, D, x(9), S21, &H21E1CDE6)   ' 25
    D = GG(D, A, b, C, x(14), S22, &HC33707D6)  ' 26
    C = GG(C, D, A, b, x(3), S23, &HF4D50D87)   ' 27
    b = GG(b, C, D, A, x(8), S24, &H455A14ED)   ' 28
    A = GG(A, b, C, D, x(13), S21, &HA9E3E905)  ' 29
    D = GG(D, A, b, C, x(2), S22, &HFCEFA3F8)   ' 30
    C = GG(C, D, A, b, x(7), S23, &H676F02D9)   ' 31
    b = GG(b, C, D, A, x(12), S24, &H8D2A4C8A)  ' 32

    ' Round 3
    A = HH(A, b, C, D, x(5), S31, &HFFFA3942)   ' 33
    D = HH(D, A, b, C, x(8), S32, &H8771F681)   ' 34
    C = HH(C, D, A, b, x(11), S33, &H6D9D6122)  ' 35
    b = HH(b, C, D, A, x(14), S34, &HFDE5380C)  ' 36
    A = HH(A, b, C, D, x(1), S31, &HA4BEEA44)   ' 37
    D = HH(D, A, b, C, x(4), S32, &H4BDECFA9)   ' 38
    C = HH(C, D, A, b, x(7), S33, &HF6BB4B60)   ' 39
    b = HH(b, C, D, A, x(10), S34, &HBEBFBC70)  ' 40
    A = HH(A, b, C, D, x(13), S31, &H289B7EC6)  ' 41
    D = HH(D, A, b, C, x(0), S32, &HEAA127FA)   ' 42
    C = HH(C, D, A, b, x(3), S33, &HD4EF3085)   ' 43
    b = HH(b, C, D, A, x(6), S34, &H4881D05)    ' 44
    A = HH(A, b, C, D, x(9), S31, &HD9D4D039)   ' 45
    D = HH(D, A, b, C, x(12), S32, &HE6DB99E5)  ' 46
    C = HH(C, D, A, b, x(15), S33, &H1FA27CF8)  ' 47
    b = HH(b, C, D, A, x(2), S34, &HC4AC5665)   ' 48

    ' Round 4
    A = II(A, b, C, D, x(0), S41, &HF4292244)   ' 49
    D = II(D, A, b, C, x(7), S42, &H432AFF97)   ' 50
    C = II(C, D, A, b, x(14), S43, &HAB9423A7)  ' 51
    b = II(b, C, D, A, x(5), S44, &HFC93A039)   ' 52
    A = II(A, b, C, D, x(12), S41, &H655B59C3)  ' 53
    D = II(D, A, b, C, x(3), S42, &H8F0CCC92)   ' 54
    C = II(C, D, A, b, x(10), S43, &HFFEFF47D)  ' 55
    b = II(b, C, D, A, x(1), S44, &H85845DD1)   ' 56
    A = II(A, b, C, D, x(8), S41, &H6FA87E4F)   ' 57
    D = II(D, A, b, C, x(15), S42, &HFE2CE6E0)  ' 58
    C = II(C, D, A, b, x(6), S43, &HA3014314)   ' 59
    b = II(b, C, D, A, x(13), S44, &H4E0811A1)  ' 60
    A = II(A, b, C, D, x(4), S41, &HF7537E82)   ' 61
    D = II(D, A, b, C, x(11), S42, &HBD3AF235)  ' 62
    C = II(C, D, A, b, x(2), S43, &H2AD7D2BB)   ' 63
    b = II(b, C, D, A, x(9), S44, &HEB86D391)   ' 64

    State(0) = uwAdd(State(0), A)
    State(1) = uwAdd(State(1), b)
    State(2) = uwAdd(State(2), C)
    State(3) = uwAdd(State(3), D)

    Exit Sub

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Sub Else Resume Next

End Sub


Private Function AddRotAdd(F As Long, A As Long, b As Long, x As Long, s As Integer, ac As Long) As Long
    ' Common routine for FF, GG, HH and II
    ' #define AddRotAdd(f, a, b, c, d, x, s, ac) { \
    '  (a) += f + (x) + (UINT4)(ac); \
    '  (a) = ROTATE_LEFT ((a), (s)); \
    '  (a) += (b); \
    '  }
    On Error GoTo ErrHandler

    Dim funcName  As String:: funcName = "AddRotAdd"
    
    Dim Temp      As Long
    Temp = uwAdd(A, F)
    Temp = uwAdd(Temp, x)
    Temp = uwAdd(Temp, ac)
    Temp = uwRol(Temp, s)
    AddRotAdd = uwAdd(Temp, b)

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Private Function ff(A As Long, b As Long, C As Long, D As Long, x As Long, s As Integer, ac As Long) As Long
    ' Returns new value of a
    ' #define F(x, y, z) (((x) & (y)) | ((~x) & (z)))
    ' #define FF(a, b, c, d, x, s, ac) { \
    '  (a) += F ((b), (c), (d)) + (x) + (UINT4)(ac); \
    '  (a) = ROTATE_LEFT ((a), (s)); \
    '  (a) += (b); \
    '  }
   On Error GoTo ErrHandler

   Dim funcName  As String:: funcName = "ff"
    
    Dim t         As Long
    Dim t2        As Long
    ' F ((b), (c), (d)) = (((b) & (c)) | ((~b) & (d)))
    t = b And C
    t2 = (Not b) And D
    t = t Or t2
    ff = AddRotAdd(t, A, b, x, s, ac)


    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Private Function GG(A As Long, b As Long, C As Long, D As Long, x As Long, s As Integer, ac As Long) As Long
    ' #define G(b, c, d) (((b) & (d)) | ((c) & (~d)))
    On Error GoTo ErrHandler

    Dim funcName  As String:: funcName = "GG"

    Dim t         As Long
    Dim t2        As Long
    t = b And D
    t2 = C And (Not D)
    t = t Or t2
    GG = AddRotAdd(t, A, b, x, s, ac)

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Private Function HH(A As Long, b As Long, C As Long, D As Long, x As Long, s As Integer, ac As Long) As Long
    ' #define H(b, c, d) ((b) ^ (c) ^ (d))
    On Error GoTo ErrHandler

    Dim funcName  As String:: funcName = "HH"

    Dim t         As Long
    t = b Xor C Xor D
    HH = AddRotAdd(t, A, b, x, s, ac)

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Private Function II(A As Long, b As Long, C As Long, D As Long, x As Long, s As Integer, ac As Long) As Long
    ' #define I(b, c, d) ((c) ^ ((b) | (~d)))
    On Error GoTo ErrHandler

    Dim funcName  As String:: funcName = "II"
    Dim t         As Long
    t = b Or (Not D)
    t = C Xor t
    II = AddRotAdd(t, A, b, x, s, ac)


    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
' Unsigned 32-bit word functions suitable for VB/VBA
Private Function uwRol(w As Long, s As Integer) As Long
    ' Return 32-bit word w rotated left by s bits
    ' avoiding problem with VB sign bit
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "uwRol"

    Dim i         As Integer
    Dim t         As Long

    uwRol = w
    For i = 1 To s
        t = uwRol And &H3FFFFFFF
        t = t * 2
        If (uwRol And &H40000000) <> 0 Then
            t = t Or &H80000000
        End If
        If (uwRol And &H80000000) <> 0 Then
            t = t Or &H1
        End If
        uwRol = t
    Next


    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Private Function uwJoin(A As Byte, b As Byte, C As Byte, D As Byte) As Long
    ' Join 4 x 8-bit bytes into one 32-bit word a.b.c.d
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "uwJoin"

    uwJoin = ((A And &H7F) * &H1000000) Or (b * &H10000) Or (CLng(C) * &H100) Or D
    If A And &H80 Then
        uwJoin = uwJoin Or &H80000000
    End If

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Private Sub uwSplit(ByVal w As Long, A As Byte, b As Byte, C As Byte, D As Byte)
    ' Split 32-bit word w into 4 x 8-bit bytes
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "uwSplit"

    A = CByte(((w And &HFF000000) \ &H1000000) And &HFF)
    b = CByte(((w And &HFF0000) \ &H10000) And &HFF)
    C = CByte(((w And &HFF00) \ &H100) And &HFF)
    D = CByte((w And &HFF) And &HFF)

    Exit Sub

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Sub Else Resume Next

End Sub
Public Function uwAdd(wordA As Long, wordB As Long) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "uwAdd"

    ' Adds words A and B avoiding overflow
    Dim myUnsigned As Double

    myUnsigned = LongToUnsigned(wordA) + LongToUnsigned(wordB)
    ' Cope with overflow
    '[2010-10-20] Changed from ">" to ">=". Thanks Loek.
    If myUnsigned >= OFFSET_4 Then
        myUnsigned = myUnsigned - OFFSET_4
    End If
    uwAdd = UnsignedToLong(myUnsigned)

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


'****************************************************
' These two functions from Microsoft Article Q189323
' "HOWTO: convert between Signed and Unsigned Numbers"
Private Function UnsignedToLong(Value As Double) As Long
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "UnsignedToLong"

    If Value < 0 Or Value >= OFFSET_4 Then Error 6    ' Overflow
    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Private Function LongToUnsigned(Value As Long) As Double
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "LongToUnsigned"

    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


