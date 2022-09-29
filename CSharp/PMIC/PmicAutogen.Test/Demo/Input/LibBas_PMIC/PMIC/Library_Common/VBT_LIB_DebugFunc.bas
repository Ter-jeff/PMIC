Attribute VB_Name = "VBT_LIB_DebugFunc"
Option Explicit


Public Const CSVcomma = ","
'20170814 evans.lo: for CT request
Dim gbGlobalAddrMap As Boolean
Dim glGlobalAddrMap() As Long
Dim glDictGlobalAddrMap() As String

'20180503 evans.lo: for Avus field mask request
Dim gsAHBFieldName() As String
Dim glAHBFieldMask() As Long

'20171027 evans: for Reg debug check
Const gbRegDumpOfflineCheck As Boolean = False    'It's for offline check

Private Site      As Variant
Private iSite     As Variant
Private g_Site    As Variant

Private Enum GLOBAL_ADDR_MAP_INDEX
    G_START_ROW = 2
    G_REG_NAME = 5
    G_REG_ADDR = 3
    G_REG_FIELD = 6
    G_REG_FIELD_Width = 7
    G_REF_FIELD_Offset = 8
End Enum

Private Enum REG_STATUS_INDEX
    S_REG_SITE = 1
    S_REG_NAME = 2
    S_REG_ADDR = 3
    S_REG_BEFORE = 4
    S_REG_AFTER = 5
    S_REG_CHECK = 6
End Enum


Private Type REG_STATUS_EXPORT
    RegName()     As String
    RegAddr()     As String
    BefData()     As String
    AftData()     As String
    ExportToFile() As String
End Type

Public Type PIN_CONDI_EXPORT
    PinName()     As String
    ChannelType() As String
    GetStatus()   As Long
    GetVoltage()  As Double
    BleederResistor() As String
    ExportToFile() As String
    GetGateStaus() As String
    PPMUcon()     As String
End Type

Public Const gs_GetStatus = "Conn"
Public Const gs_GetVoltage = "Volt"
Public Const gs_BleederResistor = "BleederResistor"
Public Const gs_GetGateStaus = "GetGateStaus"
Public Const gs_PPMUcon = "PPMUcon"



Private Type REG_FILE_READ
    READDATA()    As String
End Type

Private FilegReadBySite() As REG_FILE_READ
Private PinStatusFilegReadBySite() As REG_FILE_READ

Public ExportPinCondiDiffBySite_before() As PIN_CONDI_EXPORT
Public ExportPinCondiDiffBySite_After() As PIN_CONDI_EXPORT


Private ExportRegStatusBySite() As REG_STATUS_EXPORT
Private ExportRegStatusDiffBySite() As REG_STATUS_EXPORT

Private Const DiffCheck = "Y"

Private Const gS_REGCHECKFileDir = ".\REGCHECK\"

Public g_DigiState_Stru As DigiState_Stru

Public Type DigiState_Stru
    State_Before() As String
    State_After() As String
End Type

'20190311 JY: export reg data and Pin status to TXT file
Public Function GetStatusAllDiff(start_reg As String, end_reg As String, RegCheck As REG_DATA, Optional b_GetRegData As Boolean = True, _
                                 Optional b_GetPinCondi As Boolean = True, Optional FileName As String = "StatusAllDiff", Optional ExportToFile As Boolean = True)


    If b_GetRegData = False And b_GetPinCondi = False Then
        TheExec.Datalog.WriteComment "Neither RegData nor PinCondi will be compared !!!"
        Exit Function
    End If

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetStatusAllDiff"
    
    Dim SiteObject As Object
    Dim St_WriteTxtFile As TextStream
    Dim mS_HEADER As String
    Dim mS_FILETYPE As String:: mS_FILETYPE = ".txt"
    Dim mS_FileName As String
    Dim mS_File   As String
    Dim fs        As New FileSystemObject
    Dim St_ReadTxtFile As TextStream
    Dim read_count As Long:: read_count = 0
    Dim Index     As Long:: Index = 0
    Dim AllPinCount As Long:: AllPinCount = 0
    Dim AfterPinCount As Long:: AfterPinCount = 0
    Dim DiffPinCount As Long:: DiffPinCount = 0
    Dim ReadExportPinStatusFileArray() As String
    Dim PinStatusArray() As String

    If gbRegDumpOfflineCheck = True Then
        Set SiteObject = TheExec.Sites.Existing
    Else
        Set SiteObject = TheExec.Sites.Selected
    End If



    If RegCheck = REG_DATA_BEFORE Then

        If b_GetRegData = True Then GetDataByRegName start_reg, end_reg, REG_DATA_BEFORE, , False, True, True
        If b_GetPinCondi = True Then
            LIB_GetSetupSummary CHECK_DATA_BEFORE, True
            mS_FileName = FileName + "_before" + mS_FILETYPE
            mS_File = gS_REGCHECKFileDir + mS_FileName
            Set St_WriteTxtFile = fs.OpenTextFile(mS_File, ForWriting, True)
            'St_WriteTxtFile.WriteLine mS_HEADER
            For Each iSite In SiteObject
                For Index = 0 To UBound(ExportPinCondiDiffBySite_before(iSite).ExportToFile)
                    St_WriteTxtFile.WriteLine ExportPinCondiDiffBySite_before(iSite).ExportToFile(Index)
                Next Index
                Exit For
            Next iSite
            St_WriteTxtFile.Close
        End If

    ElseIf RegCheck = REG_DATA_AFTER Then


        ReDim PinStatusFilegReadBySite(TheExec.Sites.Existing.Count - 1)    'offline check
        ReDim ExportPinCondiDiffBySite_before(TheExec.Sites.Existing.Count - 1)
        ReDim ExportPinCondiDiffBySite_After(TheExec.Sites.Existing.Count - 1)

        If b_GetRegData = True Then GetDataByRegName start_reg, end_reg, REG_DATA_AFTER, , False, True, True
        If b_GetPinCondi = True Then
            LIB_GetSetupSummary CHECK_DATA_After, True
            'Read back previos pin status data
            mS_FileName = FileName + "_before" + mS_FILETYPE
            mS_File = gS_REGCHECKFileDir + mS_FileName
            Set St_ReadTxtFile = fs.OpenTextFile(mS_File, ForReading, True)
            For Each g_Site In TheExec.Sites
                Do While Not St_ReadTxtFile.AtEndOfStream
                    ReDim Preserve ReadExportPinStatusFileArray(read_count)
                    ReadExportPinStatusFileArray(read_count) = St_ReadTxtFile.ReadLine
                    PinStatusArray = Split(ReadExportPinStatusFileArray(read_count), CSVcomma)
                    ReDim Preserve PinStatusArray(6)
                    ReDim Preserve ExportPinCondiDiffBySite_before(g_Site).PinName(read_count)
                    ReDim Preserve ExportPinCondiDiffBySite_before(g_Site).ChannelType(read_count)
                    ReDim Preserve ExportPinCondiDiffBySite_before(g_Site).GetStatus(read_count)
                    ReDim Preserve ExportPinCondiDiffBySite_before(g_Site).GetVoltage(read_count)
                    ReDim Preserve ExportPinCondiDiffBySite_before(g_Site).BleederResistor(read_count)
                    ReDim Preserve ExportPinCondiDiffBySite_before(g_Site).GetGateStaus(read_count)
                    ReDim Preserve ExportPinCondiDiffBySite_before(g_Site).PPMUcon(read_count)
                    ExportPinCondiDiffBySite_before(g_Site).PinName(read_count) = PinStatusArray(0)    'PinName
                    ExportPinCondiDiffBySite_before(g_Site).ChannelType(read_count) = PinStatusArray(1)    'ChannelType
                    ExportPinCondiDiffBySite_before(g_Site).GetStatus(read_count) = PinStatusArray(2)    'GetStatus
                    ExportPinCondiDiffBySite_before(g_Site).GetVoltage(read_count) = PinStatusArray(3)    'GetVoltage
                    ExportPinCondiDiffBySite_before(g_Site).BleederResistor(read_count) = PinStatusArray(4)    'GetBleedeResistor
                    ExportPinCondiDiffBySite_before(g_Site).GetGateStaus(read_count) = PinStatusArray(5)    'GetGateStaus
                    ExportPinCondiDiffBySite_before(g_Site).PPMUcon(read_count) = PinStatusArray(6)    'PPMUcon
                    read_count = read_count + 1
                Loop
                Exit For
            Next g_Site
            St_ReadTxtFile.Close


            For Each g_Site In TheExec.Sites

                For AllPinCount = 0 To UBound(ExportPinCondiDiffBySite_before(g_Site).PinName)

                    If ExportPinCondiDiffBySite_After(g_Site).PinName(AfterPinCount) = "All Pin Condi are identical" Then Exit For
                    If ExportPinCondiDiffBySite_After(g_Site).PinName(AfterPinCount) = ExportPinCondiDiffBySite_before(g_Site).PinName(AllPinCount) Then

                        If ExportPinCondiDiffBySite_After(g_Site).GetStatus(AfterPinCount) <> ExportPinCondiDiffBySite_before(g_Site).GetStatus(AllPinCount) Then
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).ExportToFile(DiffPinCount)
                            ExportPinCondiDiffBySite_After(g_Site).ExportToFile(DiffPinCount) = CombineCodeForBeforeStatus(gs_GetStatus, g_Site, AllPinCount, AfterPinCount)
                            DiffPinCount = DiffPinCount + 1
                        End If
                        If ExportPinCondiDiffBySite_After(g_Site).GetVoltage(AfterPinCount) <> ExportPinCondiDiffBySite_before(g_Site).GetVoltage(AllPinCount) Then
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).ExportToFile(DiffPinCount)
                            ExportPinCondiDiffBySite_After(g_Site).ExportToFile(DiffPinCount) = CombineCodeForBeforeStatus(gs_GetVoltage, g_Site, AllPinCount, AfterPinCount)
                            DiffPinCount = DiffPinCount + 1
                        End If
                        If ExportPinCondiDiffBySite_After(g_Site).BleederResistor(AfterPinCount) <> ExportPinCondiDiffBySite_before(g_Site).BleederResistor(AllPinCount) Then
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).ExportToFile(DiffPinCount)
                            ExportPinCondiDiffBySite_After(g_Site).ExportToFile(DiffPinCount) = CombineCodeForBeforeStatus(gs_BleederResistor, g_Site, AllPinCount, AfterPinCount)
                            DiffPinCount = DiffPinCount + 1
                        End If
                        If ExportPinCondiDiffBySite_After(g_Site).GetGateStaus(AfterPinCount) <> ExportPinCondiDiffBySite_before(g_Site).GetGateStaus(AllPinCount) Then
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).ExportToFile(DiffPinCount)
                            ExportPinCondiDiffBySite_After(g_Site).ExportToFile(DiffPinCount) = CombineCodeForBeforeStatus(gs_GetGateStaus, g_Site, AllPinCount, AfterPinCount)
                            DiffPinCount = DiffPinCount + 1
                        End If
                        If ExportPinCondiDiffBySite_After(g_Site).PPMUcon(AfterPinCount) <> ExportPinCondiDiffBySite_before(g_Site).PPMUcon(AllPinCount) Then
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).ExportToFile(DiffPinCount)
                            ExportPinCondiDiffBySite_After(g_Site).ExportToFile(DiffPinCount) = CombineCodeForBeforeStatus(gs_PPMUcon, g_Site, AllPinCount, AfterPinCount)
                            DiffPinCount = DiffPinCount + 1
                        End If
                        AfterPinCount = AfterPinCount + 1    'to compare the next difference
                    End If
                    If UBound(ExportPinCondiDiffBySite_After(g_Site).PinName) + 1 = AfterPinCount Then Exit For    'finish comparison
                Next AllPinCount
                Exit For    'All site status are same
            Next g_Site
        End If    'If b_GetPinCondi = True



        '******************************
        'create StatusAllDiff.csv file
        '******************************
        mS_FileName = FileName + mS_FILETYPE
        mS_File = gS_REGCHECKFileDir + mS_FileName

        If InStr(mS_FileName, FileName) = 0 And Len(FileName) > 0 Then
            mS_FileName = FileName
            If InStr(FileName, mS_FILETYPE) = 0 Then
                mS_FileName = mS_FileName + mS_FILETYPE
            End If
        End If

        mS_File = gS_REGCHECKFileDir + mS_FileName
        mS_HEADER = "Category" + vbTab + "Site" + vbTab + "Reg/Pins" + vbTab + "Type" + vbTab + "Before" + vbTab + "After" + vbTab + "CodeForBeforeCondi"

        Call File_CheckAndCreateFolder(gS_REGCHECKFileDir)
        Call File_CreateAFile(mS_File, mS_HEADER)



        '******************************
        'write StatusAllDiff.csv file
        '******************************
        Set St_WriteTxtFile = fs.OpenTextFile(mS_File, ForWriting, True)
        St_WriteTxtFile.WriteLine mS_HEADER
        If b_GetRegData = True Then
            For Each iSite In SiteObject
                If ExportRegStatusDiffBySite(iSite).ExportToFile(0) <> "All Reg are identical" Then
                    For Index = 0 To UBound(ExportRegStatusDiffBySite(iSite).ExportToFile)
                        St_WriteTxtFile.WriteLine ExportRegStatusDiffBySite(iSite).ExportToFile(Index)
                    Next Index
                Else
                    TheExec.Datalog.WriteComment "All Registers are identical"
                End If
            Next iSite
        End If
        If b_GetPinCondi = True Then
            For Each iSite In SiteObject
                If ExportPinCondiDiffBySite_After(g_Site).PinName(0) <> "All Pin Condi are identical" Then
                    For Index = 0 To UBound(ExportPinCondiDiffBySite_After(iSite).ExportToFile)
                        St_WriteTxtFile.WriteLine ExportPinCondiDiffBySite_After(iSite).ExportToFile(Index)
                    Next Index
                Else
                    TheExec.Datalog.WriteComment "All Pins Condition are identical"
                End If
                Exit For    'All site status are same
            Next iSite
        End If
        St_WriteTxtFile.Close

    End If    ' End of RegCheck


    Exit Function

ErrHandler:
    LIB_ErrorDescription ("GetStatusAllDiff")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function




' JY 20190311
Function Dec2Bin_New(ByVal DecimalIn As Variant, _
                     Optional NumberOfBits As Variant) As String
    Dec2Bin_New = ""
    DecimalIn = Int(CDec(DecimalIn))
    Do While DecimalIn <> 0
        Dec2Bin_New = Format$(DecimalIn - 2 * Int(DecimalIn / 2)) & Dec2Bin_New
        DecimalIn = Int(DecimalIn / 2)
    Loop
    If Not IsMissing(NumberOfBits) Then
        If Len(Dec2Bin_New) > NumberOfBits Then
            Dec2Bin_New = "Error - Number exceeds specified bit size"
        Else
            Dec2Bin_New = Right$(String$(NumberOfBits, _
                                         "0") & Dec2Bin_New, NumberOfBits)
        End If
    End If
End Function


'called by GetStatusAllDiff
Public Function LIB_GetSetupSummary(Optional CHECKSTAGE As CHECK_DATA = CHECK_DATA.CHECK_DATA_BEFORE, Optional DiffPinCondiSave As Boolean = False) As Long
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "LIB_GetSetupSummary"
    Dim m_ws As Worksheet, mL_ColPinName As Long, mL_StartRow As Long, mL_StartCol As Long, mS_Temp As String
    Dim i As Long, Index As Long

    Dim CurrentChanMap As String
    Dim PinName   As String
    Dim ChannelType As String
    Dim GetVoltage As Double
    Dim GetStatus As Long
    Dim PPMUcon   As String
    Dim GetBleedeResistor As String
    Dim GetGateStaus As String

    Dim DiffCount As Long:: DiffCount = 0
    Dim bCheckResult As Boolean


    CurrentChanMap = TheExec.CurrentChanMap
    Set m_ws = ActiveWorkbook.Worksheets(CurrentChanMap)

    bCheckResult = True

    If DiffPinCondiSave = True Then    '20190307 for DiffPinCondiSave
        ReDim ExportPinCondiDiffBySite_before(TheExec.Sites.Existing.Count - 1)
        ReDim ExportPinCondiDiffBySite_After(TheExec.Sites.Existing.Count - 1)
    End If

    For Each g_Site In TheExec.Sites

        mS_Temp = FormatLog("PinName", -20) & "," & FormatLog("ChannelType", -15) & "," & _
                  FormatLog("GetStatus", -15) & "," & FormatLog("GetVoltage", -15) & "," & FormatLog("BleederResistor", -25)
        If DiffPinCondiSave = False Then
            If CHECKSTAGE = CHECK_DATA_BEFORE Then
                g_PreFlowSheetName = TheExec.Flow.CurrentFlowSheetName
                g_CurrFlowSheetName = TheExec.Flow.CurrentFlowSheetName
                TheExec.Datalog.WriteComment ""
                TheExec.Datalog.WriteComment "funcName::" & funcName
                TheExec.Datalog.WriteComment "CurrFlowSheetName = " & g_CurrFlowSheetName
                TheExec.Datalog.WriteComment mS_Temp
            ElseIf CHECKSTAGE = CHECK_DATA_After Then
                g_CurrFlowSheetName = TheExec.Flow.CurrentFlowSheetName
                TheExec.Datalog.WriteComment ""
                TheExec.Datalog.WriteComment "funcName::" & funcName
                TheExec.Datalog.WriteComment "PreFlowSheetName  = " & g_PreFlowSheetName
                TheExec.Datalog.WriteComment "CurrFlowSheetName = " & g_CurrFlowSheetName
                TheExec.Datalog.WriteComment mS_Temp & " || (Before) " & mS_Temp
            End If
        End If

        mL_StartRow = 7: Index = 0: mL_ColPinName = 2
        PinName = m_ws.Cells(mL_StartRow + Index, mL_ColPinName)
        While PinName <> ""
            If InStr(UCase(PinName), "DGS") = 0 Then
                GetStatus = 999: GetVoltage = 0: mS_Temp = "": GetBleedeResistor = ""
                ChannelType = TheExec.DataManager.ChannelType(PinName)

                Select Case ChannelType
                    Case "DCDiffMeter":
                        GetStatus = TheHdw.DCDiffMeter.Pins(PinName).Connected
                        '0 Not connected
                        '1 High sense connected
                        '2 Low sense connected
                        '3 High and low sense connected
                        '4 Low DGS connected
                        '5 High sense and low DGS connected
                    Case "DCVI":
                        GetStatus = TheHdw.DCVI.Pins(PinName).Connected
                        ' 0 Not connected         00000
                        ' 1 High force connected  00001
                        ' 2 High sense connected  00010
                        ' 4 High guard connected  00100
                        ' 8 Low force connected   01000
                        '16 Low sense connected   10000
                        GetVoltage = TheHdw.DCVI.Pins(PinName).Voltage

                        '20190311 JY
                        GetGateStaus = TheHdw.DCVI.Pins(PinName).Gate(tlDCVIGate)
                        If PinName Like "*UVI80*" Then    'Because UVI80 has GateHiZ mode
                            'TheHdw.DCVI.Pins("VDD_BUCK6_UVI80").Gate(tlDCVIGateHiZ)
                            GetGateStaus = GetGateStaus + "UVI80"
                        End If

                        Dim SlotType As String
                        SlotType = LCase(GetInstrument(PinName, 0))
                        Select Case SlotType
                            Case "hexvs": GetBleedeResistor = ""
                            Case "vhdvs": GetBleedeResistor = ""
                            Case "dc-07": GetBleedeResistor = "BleederResistor:" & TheHdw.DCVI.Pins(PinName).BleederResistor
                        End Select

                        GetBleedeResistor = Replace(GetBleedeResistor, "0", "Off")
                        GetBleedeResistor = Replace(GetBleedeResistor, "1", "On")
                        GetBleedeResistor = Replace(GetBleedeResistor, "2", "Auto")

                    Case "Utility":
                        GetStatus = TheHdw.Utility.Pins(PinName).States(tlUBStateProgrammed)
                    Case "I/O"
                        'Digital
                        GetStatus = TheHdw.Digital.Raw.Chans(PinName).IsConnected

                        'PPMU
                        PPMUcon = TheHdw.PPMU.Raw.Chans(PinName).IsConnected
                        GetGateStaus = TheHdw.PPMU.Raw.Chans(PinName).Gate

                    Case Else
                End Select

                If ChannelType = "Utility" Or ChannelType = "DCDiffMeter" Or ChannelType = "I/O" Then
                    mS_Temp = FormatLog(PinName, -20) & "," & FormatLog(ChannelType, -15) & "," & FormatLog(GetStatus, -15) & "," & FormatLog("", -15) & "," & FormatLog("", -25)
                ElseIf ChannelType = "DCVI" Then
                    mS_Temp = FormatLog(PinName, -20) & "," & FormatLog(ChannelType, -15) & "," & FormatLog(GetStatus, -15) & "," & FormatLog(Format(GetVoltage, "0.00"), -15) & "," & FormatLog(GetBleedeResistor, -25)
                    If GetStatus = 0 Then mS_Temp = FormatLog(PinName, -20) & "," & FormatLog(ChannelType, -15) & "," & FormatLog(GetStatus, -15) & "," & FormatLog("", -15) & "," & FormatLog(GetBleedeResistor, -25)
                End If

                'Compare Result::
                If CHECKSTAGE = CHECK_DATA_BEFORE Then
                    ReDim Preserve g_PreCheckData(Index): g_PreCheckData(Index) = mS_Temp
                    ReDim Preserve g_CurrCheckData(Index): g_CurrCheckData(Index) = mS_Temp

                ElseIf CHECKSTAGE = CHECK_DATA_After Then
                    ReDim Preserve g_CurrCheckData(Index): g_CurrCheckData(Index) = mS_Temp
                    ReDim Preserve g_ResultCheckData(Index): g_ResultCheckData(Index) = IIf(g_PreCheckData(Index) = g_CurrCheckData(Index), True, False)
                End If

                'Datalog
                If CHECKSTAGE = CHECK_DATA_BEFORE Then
                    If (ChannelType = "Utility" Or ChannelType = "DCDiffMeter" Or ChannelType = "I/O" Or ChannelType = "DCVI" Or ChannelType = "I/O") Then TheExec.Datalog.WriteComment mS_Temp
                    If DiffPinCondiSave = True Then
                        ReDim Preserve ExportPinCondiDiffBySite_before(g_Site).ExportToFile(DiffCount)
                        ExportPinCondiDiffBySite_before(g_Site).ExportToFile(DiffCount) = PinName + CSVcomma + ChannelType + CSVcomma + CStr(GetStatus) + CSVcomma _
                                                                                          + CStr(GetVoltage) + CSVcomma + GetBleedeResistor + CSVcomma + GetGateStaus + CSVcomma + CStr(PPMUcon)
                        DiffCount = DiffCount + 1
                    End If
                ElseIf CHECKSTAGE = CHECK_DATA_After Then
                    If g_ResultCheckData(Index) = False And (ChannelType = "Utility" Or ChannelType = "DCDiffMeter" Or ChannelType = "DCVI" Or ChannelType = "I/O") Then
                        bCheckResult = False
                        '20190307 for DiffPinCondiSave
                        If DiffPinCondiSave = True Then
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).PinName(DiffCount)
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).ChannelType(DiffCount)
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).GetStatus(DiffCount)
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).GetVoltage(DiffCount)
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).BleederResistor(DiffCount)
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).GetGateStaus(DiffCount)
                            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).PPMUcon(DiffCount)
                            ExportPinCondiDiffBySite_After(g_Site).PinName(DiffCount) = PinName
                            ExportPinCondiDiffBySite_After(g_Site).ChannelType(DiffCount) = ChannelType
                            ExportPinCondiDiffBySite_After(g_Site).GetStatus(DiffCount) = GetStatus
                            ExportPinCondiDiffBySite_After(g_Site).GetVoltage(DiffCount) = GetVoltage
                            ExportPinCondiDiffBySite_After(g_Site).BleederResistor(DiffCount) = GetBleedeResistor
                            ExportPinCondiDiffBySite_After(g_Site).GetGateStaus(DiffCount) = GetGateStaus
                            ExportPinCondiDiffBySite_After(g_Site).PPMUcon(DiffCount) = PPMUcon
                            DiffCount = DiffCount + 1
                        Else
                            TheExec.Datalog.WriteComment mS_Temp & " || (Before) " & g_PreCheckData(Index)
                        End If
                    End If
                End If

            End If
            Index = Index + 1

            PinName = m_ws.Cells(mL_StartRow + Index, mL_ColPinName)
        Wend
        If DiffCount = 0 And DiffPinCondiSave = True Then
            ReDim Preserve ExportPinCondiDiffBySite_After(g_Site).PinName(DiffCount)
            ExportPinCondiDiffBySite_After(g_Site).PinName(DiffCount) = "All Pin Condi are identical"
            TheExec.Datalog.WriteComment "All Pin Condi are identical"
        End If
        Exit For
    Next g_Site

    'Check ADG 1414 relay connection
    Dim m_ADG1414ArgList() As String
    m_ADG1414ArgList = Split(g_ADG1414ArgList, ",")
    If UBound(m_ADG1414ArgList) = -1 Then
        TheExec.AddOutput "Please SVN update VBT_DibChecker-ADG1414_CONTROL."
    Else
        For i = 0 To UBound(m_ADG1414ArgList)
            PinName = m_ADG1414ArgList(i)
            ChannelType = "ADG1414"
            GetStatus = g_ADG1414Data(i)
            mS_Temp = FormatLog(PinName, -20) & "," & FormatLog(ChannelType, -15) & "," & FormatLog(GetStatus, -15) & "," & FormatLog("", -15) & "," & FormatLog("", -25)
            'Compare Result/Datalog::
            If CHECKSTAGE = CHECK_DATA_BEFORE Then
                ReDim Preserve g_PreCheckData(Index): g_PreCheckData(Index) = mS_Temp
                ReDim Preserve g_CurrCheckData(Index): g_CurrCheckData(Index) = mS_Temp
                TheExec.Datalog.WriteComment mS_Temp
            ElseIf CHECKSTAGE = CHECK_DATA_After Then
                ReDim Preserve g_CurrCheckData(Index): g_CurrCheckData(Index) = mS_Temp
                ReDim Preserve g_ResultCheckData(Index): g_ResultCheckData(Index) = IIf(g_PreCheckData(Index) = g_CurrCheckData(Index), True, False)
                If g_ResultCheckData(Index) = False Then
                    bCheckResult = False
                    TheExec.Datalog.WriteComment mS_Temp & " || (Before) " & g_PreCheckData(Index)
                End If
            End If
            Index = Index + 1
        Next i
    End If

    If CHECKSTAGE = CHECK_DATA_After Then
        If bCheckResult = False Then TheExec.Datalog.WriteComment "**************" & funcName & ":: Check Result(Fail)**************"
    End If

    Exit Function
ErrHandler:
    LIB_ErrorDescription (funcName)
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function




'20171027 evans: exort reg data to worksheet or csv file
'20170814 evans: for CT request
Public Function GetDataByRegName(start_reg As String, end_reg As String, RegCheck As REG_DATA, Optional FileName As String = "REGCHECK.csv", _
                                 Optional ExportToSheet As Boolean = False, Optional ExportToFile As Boolean = True, Optional ExportToStatusAllDiff As Boolean = False)


    Dim SiteObject As Object
    Dim ReadExportFileArray() As String

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetDataByRegName"

    If gbRegDumpOfflineCheck = True Then
        Set SiteObject = TheExec.Sites.Existing
    Else
        Set SiteObject = TheExec.Sites.Selected
    End If

    If ExportToSheet = False And ExportToFile = False Then
        TheExec.Datalog.WriteComment "<WARNING> : The alternatives are ExportToSheet and ExportToFile!"
        Exit Function
    End If

    GlobalRegMap_Initialize

    If ExportToFile Then
        Dim mS_HEADER As String
        Dim mS_FILETYPE As String
        Dim mS_File As String
        Dim mS_FileName As String
        Dim fs    As New FileSystemObject
        Dim St_ReadTxtFile As TextStream
        Dim read_count As Long

        mS_FILETYPE = ".csv"

        mS_FileName = "REGCHECK" + mS_FILETYPE    'Join(mS_FileNameArray, "_") + mS_FILETYPE
        If InStr(mS_FileName, FileName) = 0 And Len(FileName) > 0 Then
            mS_FileName = FileName
            If InStr(FileName, mS_FILETYPE) = 0 Then
                mS_FileName = mS_FileName + mS_FILETYPE
            End If
        End If

        mS_File = gS_REGCHECKFileDir + mS_FileName
        mS_HEADER = "Site,Reg Name,Reg Addr,Before,After"

        Call File_CheckAndCreateFolder(gS_REGCHECKFileDir)
        If RegCheck = REG_DATA_BEFORE And fs.FileExists(mS_File) Then
            Call File_CreateAFile(mS_File, mS_HEADER)
        End If

        read_count = 0

        If RegCheck = REG_DATA_AFTER Then

            ReDim FilegReadBySite(TheExec.Sites.Existing.Count - 1)    'offline check

            Set St_ReadTxtFile = fs.OpenTextFile(mS_File, ForReading, True)
            St_ReadTxtFile.ReadLine    'filter header
            Do While Not St_ReadTxtFile.AtEndOfStream
                ReDim Preserve ReadExportFileArray(read_count)
                ReadExportFileArray(read_count) = St_ReadTxtFile.ReadLine
                read_count = read_count + 1
            Loop
            St_ReadTxtFile.Close
            Call ReSortReadRegData(ReadExportFileArray, FilegReadBySite)
        End If
    End If

    If gbGlobalAddrMap = True Then
        Dim Index As Long
        Dim RegData As New SiteLong
        Dim iSite As Variant
        Dim RegStatusCheckSheet As Object
        Dim Count As Long:: Count = 0
        Dim DiffCount As New SiteLong
        Dim start_index, end_index As Long
        Dim start_addr As Long, end_addr As Long

        start_addr = -1
        end_addr = -1

        For Index = 0 To UBound(glDictGlobalAddrMap)
            If glDictGlobalAddrMap(Index) = start_reg Then start_addr = glGlobalAddrMap(Index)
            If glDictGlobalAddrMap(Index) = end_reg Then end_addr = glGlobalAddrMap(Index)
        Next Index

        If start_addr = -1 Or end_addr = -1 Then
            TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Register Name(it should be from GlobalAddressMap worksheet!)"
            Exit Function
        End If

        If end_addr < start_addr Then
            TheExec.Datalog.WriteComment "<ERROR> : The end register address should larger than start register address"
            Exit Function
        End If

        start_index = -1
        end_index = -1

        For Index = 0 To UBound(glGlobalAddrMap)
            If glGlobalAddrMap(Index) = start_addr Then start_index = Index
            If glGlobalAddrMap(Index) = end_addr Then end_index = Index
        Next Index

        If start_index = -1 Or end_index = -1 Then
            TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Address(it should be from GlobalAddressMap worksheet!)"
            Exit Function
        End If

        If (end_index - start_index) > &HFF Then
            TheExec.Datalog.WriteComment "<WARNING> : Dump over 100 registers only export to CSV file!"
            ExportToSheet = False
            ExportToFile = True
        End If

        Worksheets("REG_STATUS_CHECK").Activate
        Set RegStatusCheckSheet = ThisWorkbook.Sheets("REG_STATUS_CHECK")

        If RegCheck = REG_DATA_BEFORE Then
            If Len(RegStatusCheckSheet.Cells(2, S_REG_BEFORE).Value) > 0 Then
                RegStatusCheckSheet.UsedRange.ClearContents
            End If
        End If


        ReDim ExportRegStatusBySite(TheExec.Sites.Existing.Count - 1)    'offline check
        ReDim ExportRegStatusDiffBySite(TheExec.Sites.Existing.Count - 1)



        For Index = 0 To UBound(glGlobalAddrMap)
            If glGlobalAddrMap(Index) >= start_addr And glGlobalAddrMap(Index) <= end_addr Then
                Call AHB_READ(glGlobalAddrMap(Index), RegData)

                TheHdw.Wait 0.05
                '''                TheExec.Datalog.WriteComment "RegAddr:" & glDictGlobalAddrMap(Index)
                For Each iSite In SiteObject
                    If TheExec.TesterMode = testModeOffline Then RegData = 1    'offline check
                    If glGlobalAddrMap(Index) = start_addr Then DiffCount(iSite) = 0
                    ReDim Preserve ExportRegStatusBySite(iSite).RegName(Count)
                    ExportRegStatusBySite(iSite).RegName(Count) = glDictGlobalAddrMap(Index)

                    ReDim Preserve ExportRegStatusBySite(iSite).RegAddr(Count)
                    ExportRegStatusBySite(iSite).RegAddr(Count) = "0x" & CStr(Hex(glGlobalAddrMap(Index)))

                    If RegCheck = REG_DATA_BEFORE Then
                        ReDim Preserve ExportRegStatusBySite(iSite).BefData(Count)
                        ExportRegStatusBySite(iSite).BefData(Count) = CStr(Hex(RegData(iSite)))

                        If ExportToFile = True Then
                            ReDim Preserve ExportRegStatusBySite(iSite).ExportToFile(Count)
                            ExportRegStatusBySite(iSite).ExportToFile(Count) = ComposeExportString(iSite, ExportRegStatusBySite(iSite), Count)
                        End If

                    End If

                    If RegCheck = REG_DATA_AFTER Then
                        ReDim Preserve ExportRegStatusBySite(iSite).AftData(Count)
                        ExportRegStatusBySite(iSite).AftData(Count) = CStr(Hex(RegData(iSite)))

                        If ExportToFile = True Then
                            If CompareExportString(ReadExportFileArray(Count), glDictGlobalAddrMap(Index)) = False Then
                                TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Address(it should be same as Previous Address Setup!)"
                                Exit Function
                            Else
                                ReDim Preserve ExportRegStatusBySite(iSite).ExportToFile(Count)
                                ExportRegStatusBySite(iSite).ExportToFile(Count) = FilegReadBySite(iSite).READDATA(Count) + ExportRegStatusBySite(iSite).AftData(Count) _
                                                                                   + ComposeRegData(FilegReadBySite(iSite).READDATA(Count), ExportRegStatusBySite(iSite).AftData(Count))
                                If Right(ComposeRegData(FilegReadBySite(iSite).READDATA(Count), ExportRegStatusBySite(iSite).AftData(Count)), Len(DiffCheck)) = DiffCheck And ExportToStatusAllDiff = True And ExportRegStatusBySite(iSite).RegName(Count) <> "" Then
                                    ReDim Preserve ExportRegStatusDiffBySite(iSite).ExportToFile(DiffCount(iSite))
                                    ExportRegStatusDiffBySite(iSite).ExportToFile(DiffCount(iSite)) = CombineCodeForBeforeValue("Register," + FilegReadBySite(iSite).READDATA(Count) + ExportRegStatusBySite(iSite).AftData(Count))
                                    DiffCount(iSite) = DiffCount(iSite) + 1
                                End If
                            End If
                        End If

                    End If

                Next iSite
                Count = Count + 1
            End If
        Next Index
        If RegCheck = REG_DATA_AFTER Then
            For Each iSite In TheExec.Sites
                If DiffCount(iSite) = 0 Then
                    ReDim Preserve ExportRegStatusDiffBySite(iSite).ExportToFile(DiffCount(iSite))
                    ExportRegStatusDiffBySite(iSite).ExportToFile(DiffCount(iSite)) = "All Reg are identical"
                End If
            Next iSite
        End If
        If ExportToSheet Then
            Index = 0
            Count = 2
            RegStatusCheckSheet.Cells(Index + 1, S_REG_SITE).Value = "Site"
            RegStatusCheckSheet.Cells(Index + 1, S_REG_NAME).Value = "Reg Name"
            RegStatusCheckSheet.Cells(Index + 1, S_REG_ADDR).Value = "Reg Addr"
            RegStatusCheckSheet.Cells(Index + 1, S_REG_BEFORE).Value = "Before"
            RegStatusCheckSheet.Cells(Index + 1, S_REG_AFTER).Value = "After"
            For Each iSite In SiteObject
                For Index = 0 To UBound(ExportRegStatusBySite(iSite).RegName)
                    RegStatusCheckSheet.Cells(Index + Count, S_REG_SITE).Value = CStr(iSite)
                    RegStatusCheckSheet.Cells(Index + Count, S_REG_NAME).Value = ExportRegStatusBySite(iSite).RegName(Index)
                    RegStatusCheckSheet.Cells(Index + Count, S_REG_ADDR).Value = "0x" & ExportRegStatusBySite(iSite).RegAddr(Index)
                    If RegCheck = REG_DATA_BEFORE Then RegStatusCheckSheet.Cells(Index + Count, S_REG_BEFORE).Value = ExportRegStatusBySite(iSite).BefData(Index)
                    If RegCheck = REG_DATA_AFTER Then
                        RegStatusCheckSheet.Cells(Index + Count, S_REG_AFTER).Value = ExportRegStatusBySite(iSite).AftData(Index)
                        If RegStatusCheckSheet.Cells(Index + Count, S_REG_AFTER).Value <> RegStatusCheckSheet.Cells(Index + Count, S_REG_BEFORE).Value Then
                            RegStatusCheckSheet.Cells(Index + Count, S_REG_CHECK).Value = DiffCheck
                        End If
                    End If
                Next Index
                Count = Count + UBound(ExportRegStatusBySite(iSite).RegName) + 1
            Next iSite
        End If

        If ExportToFile Then
            Dim St_WriteTxtFile As TextStream

            Set St_WriteTxtFile = fs.OpenTextFile(mS_File, ForWriting, True)
            St_WriteTxtFile.WriteLine mS_HEADER
            For Each iSite In SiteObject
                For Index = 0 To UBound(ExportRegStatusBySite(iSite).ExportToFile)
                    St_WriteTxtFile.WriteLine ExportRegStatusBySite(iSite).ExportToFile(Index)
                Next Index
            Next iSite

            St_WriteTxtFile.Close

        End If

    End If

    TheExec.Datalog.WriteComment "<GetDataByRegName> : Register dump is completed!"

    Exit Function

ErrHandler:
    LIB_ErrorDescription ("GetDataByRegName")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function



'20190308 JY
'20190314 JY --> Add PPMU
Private Function CombineCodeForBeforeStatus(ReadFileString As String, whichsite As Variant, all_pinCnt As Long, Af_pinCnt As Long) As String

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "CombineCodeForBeforeStatus"
    
    Dim TmpArray() As String
    Dim CodeForBeforeStatus As String:: CodeForBeforeStatus = ""
    Dim ConnEachState() As String
    Dim sDCVIConnState As String:: sDCVIConnState = ""
    Dim iDCVIConnState As Integer:: iDCVIConnState = 0
    Dim i         As Integer:: i = 0

    CombineCodeForBeforeStatus = ""


    Select Case ExportPinCondiDiffBySite_before(whichsite).ChannelType(all_pinCnt)

        Case "DCDiffMeter":
            CodeForBeforeStatus = "TheHdw.DCDiffMeter.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Connect "

            Select Case ExportPinCondiDiffBySite_before(whichsite).GetStatus(all_pinCnt)
                Case 0:
                    CodeForBeforeStatus = CodeForBeforeStatus + "tlDCVSConnectNone"
                Case 1:
                    CodeForBeforeStatus = CodeForBeforeStatus + "tlDCDiffMeterHighSense"
                Case 2:
                    CodeForBeforeStatus = CodeForBeforeStatus + "tlDCDiffMeterConnectLowSense"
                Case 3:    ''tlDCDiffMeterHighSense Or tlDCDiffMeterConnectLowSense
                    CodeForBeforeStatus = CodeForBeforeStatus + "tlDCDiffMeterHighSense & tlDCDiffMeterConnectLowSense"
                Case 4:
                    CodeForBeforeStatus = CodeForBeforeStatus + "tlDCDiffMeterConnectLowDGS"
                Case 5:

                Case Else: Exit Function   'Stop 'replace Stop with Exit Function 2019_1213
            End Select

        Case "DCVI":

            Select Case ReadFileString

                Case "Conn":
                    sDCVIConnState = Dec2Bin_New(CStr(ExportPinCondiDiffBySite_before(whichsite).GetStatus(all_pinCnt)), 5)
                    ReDim ConnEachState(Len(sDCVIConnState) - 1)
                    For i = 1 To Len(sDCVIConnState)
                        ConnEachState(i - 1) = Mid$(sDCVIConnState, i, 1)
                        iDCVIConnState = iDCVIConnState + CInt(ConnEachState(i - 1))
                    Next i
                    ' 0 Not connected         00000 tlDCVIConnectNone
                    ' 1 High force connected  00001 tlDCVIConnectHighForce
                    ' 2 High sense connected  00010 tlDCVIConnectHighSense
                    ' 4 High guard connected  00100 tlDCVIConnectHighGuard
                    ' 8 Low force connected   01000 tlDCVIConnectLowForce
                    '16 Low sense connected   10000 tlDCVIConnectLowSense
                    If iDCVIConnState = 0 Then
                        CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Disconnect"
                    Else
                        If ConnEachState(0) = 1 Then
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Connect " + "tlDCVIConnectHighForce"
                        Else
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Disconnect " + "tlDCVIConnectHighForce"
                        End If
                        If ConnEachState(1) = 1 Then
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Connect " + "tlDCVIConnectHighSense"
                        Else
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Disconnect " + "tlDCVIConnectHighSense"
                        End If
                        If ConnEachState(2) = 1 Then
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Connect " + "tlDCVIConnectHighGuard"
                        Else
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Disconnect " + "tlDCVIConnectHighGuard"
                        End If
                        If ConnEachState(3) = 1 Then
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Connect " + "tlDCVIConnectLowForce"
                        Else
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Disconnect " + "tlDCVIConnectLowForce"
                        End If
                        If ConnEachState(4) = 1 Then
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Connect " + "tlDCVIConnectLowSense"
                        Else
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Disconnect " + "tlDCVIConnectLowSense"
                        End If
                    End If

                Case "Volt":
                    CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Voltage = " _
                                          + CStr(Round(ExportPinCondiDiffBySite_before(whichsite).GetVoltage(all_pinCnt), 2))
                Case "BleederResistor":
                    CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").BleederResistor ="
                    Select Case ExportPinCondiDiffBySite_before(whichsite).BleederResistor(all_pinCnt)

                        Case "BleederResistor:Off":
                            CodeForBeforeStatus = CodeForBeforeStatus + "tlDCVIBleederResistorOff"
                        Case "BleederResistor:On":
                            CodeForBeforeStatus = CodeForBeforeStatus + "tlDCVIBleederResistorOn"
                        Case "BleederResistor:Auto":
                            CodeForBeforeStatus = CodeForBeforeStatus + "tlDCVIBleederResistorAuto"
                    End Select

                Case "GetGateStaus":
                    If ExportPinCondiDiffBySite_before(whichsite).GetGateStaus(all_pinCnt) Like "*True*" Then

                        CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Gate=True"
                    Else
                        If ExportPinCondiDiffBySite_before(whichsite).GetGateStaus(all_pinCnt) Like "*UVI80*" Then
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Gate=False"
                            CodeForBeforeStatus = CodeForBeforeStatus + vbTab + "or" + vbTab + _
                                                  "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Gate(tlDCVIGateHiZ)=False" + vbTab + _
                                                  "Please use TDE to reconfirm gate status since we can't tell it's GateOff or GateOffHiz"
                        Else
                            CodeForBeforeStatus = "TheHdw.DCVI.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Gate=False"
                        End If
                    End If
            End Select


        Case "Utility":
            'TheHdw.Utility.Pins(PinName).States(tlUBStateProgrammed)

            Select Case ExportPinCondiDiffBySite_before(whichsite).GetStatus(all_pinCnt)
                Case 0:
                    CodeForBeforeStatus = "TheHdw.Utility.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").state = tlUtilBitOff"
                Case 1:
                    CodeForBeforeStatus = "TheHdw.Utility.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").state = tlUtilBitOn"
            End Select

        Case "I/O"
            If ReadFileString = "Conn" Then
                Select Case ExportPinCondiDiffBySite_before(whichsite).GetStatus(all_pinCnt)
                    Case 0:
                        CodeForBeforeStatus = "TheHdw.Digital.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Connect"
                    Case -1:
                        CodeForBeforeStatus = "TheHdw.Digital.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Disconnect"
                End Select
            End If
            If ReadFileString = "GetGateStaus" Then
                Select Case ExportPinCondiDiffBySite_before(whichsite).GetGateStaus(all_pinCnt)
                    Case "-1":
                        CodeForBeforeStatus = "TheHdw.PPMU.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Gate =tlOn"
                    Case "0":
                        CodeForBeforeStatus = "TheHdw.PPMU.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Gate =tlOff"
                End Select
            End If
            If ReadFileString = "PPMUcon" Then
                Select Case ExportPinCondiDiffBySite_before(whichsite).PPMUcon(all_pinCnt)
                    Case "True":
                        CodeForBeforeStatus = "TheHdw.PPMU.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Connect"
                    Case "False":
                        CodeForBeforeStatus = "TheHdw.PPMU.Pins(" + Chr(34) + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + Chr(34) + ").Disconnect"
                End Select
            End If
    End Select

    'Export File
    Select Case ReadFileString
        Case gs_GetStatus:
            CombineCodeForBeforeStatus = ExportPinCondiDiffBySite_before(whichsite).ChannelType(all_pinCnt) + vbTab _
                                         + "" + vbTab + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + vbTab + ReadFileString + vbTab _
                                         + CStr(ExportPinCondiDiffBySite_before(whichsite).GetStatus(all_pinCnt)) + vbTab + _
                                         CStr(ExportPinCondiDiffBySite_After(whichsite).GetStatus(Af_pinCnt)) + vbTab + CodeForBeforeStatus
        Case gs_GetVoltage:
            CombineCodeForBeforeStatus = ExportPinCondiDiffBySite_before(whichsite).ChannelType(all_pinCnt) + vbTab _
                                         + "" + vbTab + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + vbTab + ReadFileString + vbTab _
                                         + CStr(Round(ExportPinCondiDiffBySite_before(whichsite).GetVoltage(all_pinCnt), 2)) + vbTab + _
                                         CStr(Round(ExportPinCondiDiffBySite_After(whichsite).GetVoltage(Af_pinCnt), 2)) + vbTab + CodeForBeforeStatus
        Case gs_BleederResistor:
            CombineCodeForBeforeStatus = ExportPinCondiDiffBySite_before(whichsite).ChannelType(all_pinCnt) + vbTab _
                                         + "" + vbTab + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + vbTab + ReadFileString + vbTab _
                                         + CStr(ExportPinCondiDiffBySite_before(whichsite).BleederResistor(all_pinCnt)) + vbTab + _
                                         CStr(ExportPinCondiDiffBySite_After(whichsite).BleederResistor(Af_pinCnt)) + vbTab + CodeForBeforeStatus
        Case gs_GetGateStaus:
            CombineCodeForBeforeStatus = ExportPinCondiDiffBySite_before(whichsite).ChannelType(all_pinCnt) + vbTab _
                                         + "" + vbTab + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + vbTab + ReadFileString + vbTab _
                                         + Replace(CStr(ExportPinCondiDiffBySite_before(whichsite).GetGateStaus(all_pinCnt)), "UVI80", "") + vbTab + _
                                         Replace(CStr(ExportPinCondiDiffBySite_After(whichsite).GetGateStaus(Af_pinCnt)), "UVI80", "") + vbTab + CodeForBeforeStatus
        Case gs_PPMUcon:
            CombineCodeForBeforeStatus = ExportPinCondiDiffBySite_before(whichsite).ChannelType(all_pinCnt) + vbTab _
                                         + "" + vbTab + ExportPinCondiDiffBySite_before(whichsite).PinName(all_pinCnt) + vbTab + ReadFileString + vbTab _
                                         + CStr(ExportPinCondiDiffBySite_before(whichsite).PPMUcon(all_pinCnt)) + vbTab + _
                                         CStr(ExportPinCondiDiffBySite_After(whichsite).PPMUcon(Af_pinCnt)) + vbTab + CodeForBeforeStatus
    End Select

    Exit Function
ErrHandler:
    LIB_ErrorDescription ("CombineCodeForBeforeStatus")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function

'Public Function GlobalRegMap_Initialize() As Boolean
'
'    On Error GoTo ErrHandler
'
'    Dim lastRowIndex As Range
'    Dim LastRow As Double
'    Dim RegNameIndex As Integer
'    Dim CheckRegName As String
'    Dim FieldWidth As Long
'    Dim Index As Integer
'
'    If gbGlobalAddrMap = False Then
'        Dim Row As Double
'        Dim GlobalAddressMapSheet As Object
'
'    'Find the Last Row Index of the GlobalAddressMap
'    '-------------------------------------------------------------
'
'        'Worksheets("GlobalAdressMap").Activate
'        Set GlobalAddressMapSheet = ThisWorkbook.Sheets("AHB_register_map")
'        Set lastRowIndex = GlobalAddressMapSheet.Range("A65536").End(xlUp)
'        LastRow = lastRowIndex.Row
'    '-------------------------------------------------------------
''20180503 evans.lo : For AHB address
'        RegNameIndex = 0
'        ReDim glGlobalAddrMap(RegNameIndex)
'        ReDim glDictGlobalAddrMap(RegNameIndex)
'        CheckRegName = GlobalAddressMapSheet.Cells(GLOBAL_ADDR_MAP_INDEX.G_START_ROW, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value
''20180503 evans.lo : For AHB Field Mask
'        ReDim Preserve gsAHBFieldName(LastRow - GLOBAL_ADDR_MAP_INDEX.G_START_ROW)
'        ReDim Preserve glAHBFieldMask(LastRow - GLOBAL_ADDR_MAP_INDEX.G_START_ROW)
'
'        For Row = GLOBAL_ADDR_MAP_INDEX.G_START_ROW To LastRow
''20180503 evans.lo : For AHB address
'            If CheckRegName <> GlobalAddressMapSheet.Cells(Row + 1, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value Then
'                glGlobalAddrMap(RegNameIndex) = CLng(Replace(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_ADDR).Value, "0x", "&H"))
'                glDictGlobalAddrMap(RegNameIndex) = CheckRegName 'GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value
'                CheckRegName = GlobalAddressMapSheet.Cells(Row + 1, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value
'                RegNameIndex = RegNameIndex + 1
'                If Len(CheckRegName) > 0 Then
'                    ReDim Preserve glGlobalAddrMap(RegNameIndex)
'                    ReDim Preserve glDictGlobalAddrMap(RegNameIndex)
'                End If
'            End If
''20180503 evans.lo : For AHB Field Mask
'            gsAHBFieldName(Row - GLOBAL_ADDR_MAP_INDEX.G_START_ROW) = UCase(Trim(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_FIELD).Value))
'            FieldWidth = 0
'            For Index = 0 To CLng(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_FIELD_Width)) - 1
'                FieldWidth = FieldWidth + 2 ^ Index
'            Next Index
'            FieldWidth = FieldWidth * 2 ^ CLng(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REF_FIELD_Offset))
'            glAHBFieldMask(Row - GLOBAL_ADDR_MAP_INDEX.G_START_ROW) = CLng("&H" & Mid(CStr(Hex(Not FieldWidth)), 7, 2))
'        Next Row
'        gbGlobalAddrMap = True
'    End If
'
'    GlobalRegMap_Initialize = gbGlobalAddrMap
'
'Exit Function
'
'ErrHandler:
'    LIB_ErrorDescription ("GlobalRegMap_Initialize")
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'20171030 evans: register dump and compare lib => used for export to file
Private Function ComposeExportString(Site As Variant, ExportRegStatus As REG_STATUS_EXPORT, Index As Long) As String
    Dim ExportDataArray(4) As String

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "ComposeExportString"

    ExportDataArray(0) = CStr(Site)
    ExportDataArray(1) = ExportRegStatus.RegName(Index)
    ExportDataArray(2) = ExportRegStatus.RegAddr(Index)
    ExportDataArray(3) = ExportRegStatus.BefData(Index)
    ComposeExportString = Join(ExportDataArray, ",")

    Exit Function
ErrHandler:
    LIB_ErrorDescription ("ComposeExportString")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function

Private Sub ReSortReadRegData(ReadExportFileArray() As String, FilegReadBySite() As REG_FILE_READ)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "ReSortReadRegData"

    Dim SiteObject As Object
    Dim Index     As Long
    Dim Count     As Long
    Dim iSite     As Variant

    If gbRegDumpOfflineCheck = True Then
        Set SiteObject = TheExec.Sites.Existing
    Else
        Set SiteObject = TheExec.Sites.Selected
    End If

    Count = 0

    For Each iSite In SiteObject
        For Index = 0 To UBound(ReadExportFileArray)
            If InStr(ReadExportFileArray(Index), CStr(iSite)) = 1 Then
                ReDim Preserve FilegReadBySite(iSite).READDATA(Count)
                FilegReadBySite(iSite).READDATA(Count) = ReadExportFileArray(Index)
                Count = Count + 1
            End If
        Next Index
        Count = 0
    Next iSite

    Exit Sub

ErrHandler:
    LIB_ErrorDescription ("ReSortReadRegData")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'**************************DSSC read write for debug***************************************************************************************************
'20171027 evans: exort reg data to worksheet or csv file
'20170814 evans: for CT request
Public Function GetDataByRegAddr(start_addr As Long, end_addr As Long, RegCheck As REG_DATA, _
                                 Optional ExportToSheet As Boolean = True, Optional ExportToFile As Boolean = True)

    Dim SiteObject As Object
    Dim ReadExportFileArray() As String

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetDataByRegAddr"

    If gbRegDumpOfflineCheck = True Then
        Set SiteObject = TheExec.Sites.Existing
    Else
        Set SiteObject = TheExec.Sites.Selected
    End If

    If end_addr < start_addr Then
        TheExec.Datalog.WriteComment "<ERROR> : The end address should larger than start address!"
        Exit Function
    End If

    If ExportToSheet = False And ExportToFile = False Then
        TheExec.Datalog.WriteComment "<WARNING> : The alternatives are ExportToSheet and ExportToFile!"
        Exit Function
    End If

    If end_addr < 0 Then
        TheExec.Datalog.WriteComment "<WARNING> : The end address should smaller than &H7FFF"
        Exit Function
    End If

    GlobalRegMap_Initialize

    If ExportToFile Then
        Dim mS_HEADER As String
        Dim mS_FILETYPE As String
        Dim mS_File As String
        Dim mS_FileName As String
        Dim fs    As New FileSystemObject
        Dim St_ReadTxtFile As TextStream
        Dim read_count As Long

        mS_FILETYPE = ".csv"

        mS_FileName = "REGCHECK" + mS_FILETYPE    'Join(mS_FileNameArray, "_") + mS_FILETYPE
        mS_File = gS_REGCHECKFileDir + mS_FileName
        mS_HEADER = "Site,Reg Name,Reg Addr,Before,After"

        Call File_CheckAndCreateFolder(gS_REGCHECKFileDir)
        If RegCheck = REG_DATA_BEFORE And fs.FileExists(mS_File) Then
            Call File_CreateAFile(mS_File, mS_HEADER)
        End If

        read_count = 0

        If RegCheck = REG_DATA_AFTER Then
            ReDim FilegReadBySite(TheExec.Sites.Existing.Count - 1)    'offline check

            Set St_ReadTxtFile = fs.OpenTextFile(mS_File, ForReading, True)
            St_ReadTxtFile.ReadLine    'filter header
            Do While Not St_ReadTxtFile.AtEndOfStream
                ReDim Preserve ReadExportFileArray(read_count)
                ReadExportFileArray(read_count) = St_ReadTxtFile.ReadLine
                read_count = read_count + 1
            Loop
            St_ReadTxtFile.Close
            Call ReSortReadRegData(ReadExportFileArray, FilegReadBySite)
        End If
    End If

    If gbGlobalAddrMap = True Then
        Dim Index As Long
        Dim RegData As New SiteLong
        Dim iSite As Variant
        Dim RegStatusCheckSheet As Object
        Dim Count As Long
        Dim start_index, end_index As Long

        start_index = -1
        end_index = -1

        For Index = 0 To UBound(glGlobalAddrMap)
            If glGlobalAddrMap(Index) = start_addr Then start_index = Index
            If glGlobalAddrMap(Index) = end_addr Then end_index = Index
        Next Index

        If start_index = -1 Or end_index = -1 Then
            TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Address(it should be from GlobalAddressMap worksheet!)"
            Exit Function
        End If

        If (end_index - start_index) > &HFF Then
            TheExec.Datalog.WriteComment "<WARNING> : Dump over 100 registers only export to CSV file!"
            ExportToSheet = False
            ExportToFile = True
        End If

        Worksheets("REG_STATUS_CHECK").Activate
        Set RegStatusCheckSheet = ThisWorkbook.Sheets("REG_STATUS_CHECK")

        If RegCheck = REG_DATA_BEFORE Then
            If Len(RegStatusCheckSheet.Cells(2, S_REG_BEFORE).Value) > 0 Then
                RegStatusCheckSheet.UsedRange.ClearContents
            End If
        End If

        ReDim ExportRegStatusBySite(TheExec.Sites.Existing.Count - 1)    'offline check
        Count = 0

        For Index = 0 To UBound(glGlobalAddrMap)
            If glGlobalAddrMap(Index) >= start_addr And glGlobalAddrMap(Index) <= end_addr Then
                Call AHB_READDSC(glGlobalAddrMap(Index), RegData)
                TheHdw.Wait 0.05
                '''                TheExec.Datalog.WriteComment "RegAddr:" & glDictGlobalAddrMap(Index)
                For Each iSite In SiteObject
                    ReDim Preserve ExportRegStatusBySite(iSite).RegName(Count)
                    ExportRegStatusBySite(iSite).RegName(Count) = glDictGlobalAddrMap(Index)

                    ReDim Preserve ExportRegStatusBySite(iSite).RegAddr(Count)
                    ExportRegStatusBySite(iSite).RegAddr(Count) = "0x" & CStr(Hex(glGlobalAddrMap(Index)))

                    If RegCheck = REG_DATA_BEFORE Then
                        ReDim Preserve ExportRegStatusBySite(iSite).BefData(Count)
                        ExportRegStatusBySite(iSite).BefData(Count) = CStr(Hex(RegData(iSite)))

                        If ExportToFile = True Then
                            ReDim Preserve ExportRegStatusBySite(iSite).ExportToFile(Count)
                            ExportRegStatusBySite(iSite).ExportToFile(Count) = ComposeExportString(iSite, ExportRegStatusBySite(iSite), Count)
                        End If

                    End If

                    If RegCheck = REG_DATA_AFTER Then
                        ReDim Preserve ExportRegStatusBySite(iSite).AftData(Count)
                        ExportRegStatusBySite(iSite).AftData(Count) = CStr(Hex(RegData(iSite)))

                        If ExportToFile = True Then
                            If CompareExportString(ReadExportFileArray(Count), glDictGlobalAddrMap(Index)) = False Then
                                TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Address(it should be same as Previous Address Setup!)"
                                Exit Function
                            Else
                                ReDim Preserve ExportRegStatusBySite(iSite).ExportToFile(Count)
                                ExportRegStatusBySite(iSite).ExportToFile(Count) = FilegReadBySite(iSite).READDATA(Count) + ExportRegStatusBySite(iSite).AftData(Count) _
                                                                                   + ComposeRegData(FilegReadBySite(iSite).READDATA(Count), ExportRegStatusBySite(iSite).AftData(Count))
                            End If
                        End If

                    End If

                Next iSite
                Count = Count + 1
            End If
        Next Index

        If ExportToSheet Then
            Index = 0
            Count = 2
            RegStatusCheckSheet.Cells(Index + 1, S_REG_SITE).Value = "Site"
            RegStatusCheckSheet.Cells(Index + 1, S_REG_NAME).Value = "Reg Name"
            RegStatusCheckSheet.Cells(Index + 1, S_REG_ADDR).Value = "Reg Addr"
            RegStatusCheckSheet.Cells(Index + 1, S_REG_BEFORE).Value = "Before"
            RegStatusCheckSheet.Cells(Index + 1, S_REG_AFTER).Value = "After"
            For Each iSite In SiteObject
                For Index = 0 To UBound(ExportRegStatusBySite(iSite).RegName)
                    RegStatusCheckSheet.Cells(Index + Count, S_REG_SITE).Value = CStr(iSite)
                    RegStatusCheckSheet.Cells(Index + Count, S_REG_NAME).Value = ExportRegStatusBySite(iSite).RegName(Index)
                    RegStatusCheckSheet.Cells(Index + Count, S_REG_ADDR).Value = "0x" & ExportRegStatusBySite(iSite).RegAddr(Index)
                    If RegCheck = REG_DATA_BEFORE Then RegStatusCheckSheet.Cells(Index + Count, S_REG_BEFORE).Value = ExportRegStatusBySite(iSite).BefData(Index)
                    If RegCheck = REG_DATA_AFTER Then
                        RegStatusCheckSheet.Cells(Index + Count, S_REG_AFTER).Value = ExportRegStatusBySite(iSite).AftData(Index)
                        If RegStatusCheckSheet.Cells(Index + Count, S_REG_AFTER).Value <> RegStatusCheckSheet.Cells(Index + Count, S_REG_BEFORE).Value Then
                            RegStatusCheckSheet.Cells(Index + Count, S_REG_CHECK).Value = DiffCheck
                        End If
                    End If
                Next Index
                Count = Count + UBound(ExportRegStatusBySite(iSite).RegName) + 1
            Next iSite
        End If

        If ExportToFile Then
            Dim St_WriteTxtFile As TextStream

            Set St_WriteTxtFile = fs.OpenTextFile(mS_File, ForWriting, True)
            St_WriteTxtFile.WriteLine mS_HEADER
            For Each iSite In SiteObject
                For Index = 0 To UBound(ExportRegStatusBySite(iSite).ExportToFile)
                    St_WriteTxtFile.WriteLine ExportRegStatusBySite(iSite).ExportToFile(Index)
                Next Index
            Next iSite

            St_WriteTxtFile.Close

        End If

    End If

    TheExec.Datalog.WriteComment "<GetDataByRegAddr> : Register dump is completed!"

    Exit Function


ErrHandler:
    LIB_ErrorDescription ("GetDataByRegAddr")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function

Private Function ComposeRegData(ReadFileString As String, AfterData As String) As String

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "ComposeRegData"
   
    Dim TmpArray() As String

    ComposeRegData = ""
    TmpArray = Split(ReadFileString, ",")
    If TmpArray(S_REG_BEFORE - 1) <> AfterData Then
        ComposeRegData = "," + DiffCheck
    End If
    Exit Function
ErrHandler:
    LIB_ErrorDescription ("ComposeRegData")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function

Private Function CombineCodeForBeforeValue(ReadFileString As String) As String

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "CombineCodeForBeforeValue"
    
    Dim TmpArray() As String

    CombineCodeForBeforeValue = ""
    TmpArray = Split(ReadFileString, CSVcomma)
    ReadFileString = Replace(ReadFileString, CSVcomma, vbTab)
    CombineCodeForBeforeValue = ReadFileString + vbTab + "g_RegVal = &H" + Format(TmpArray(S_REG_BEFORE), "00") + ": AHB_WRITEDSC " + TmpArray(S_REG_NAME) + ".Addr, g_RegVal"
    Exit Function
ErrHandler:
    LIB_ErrorDescription ("CombineCodeForBeforeValue")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function




''20171108 evans
Public Function GetAHBAddress(RegName As String) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetAHBAddress"
    
    Dim Index     As Integer

    GlobalRegMap_Initialize

    GetAHBAddress = -1

    For Index = 0 To UBound(glDictGlobalAddrMap)
        If UCase(glDictGlobalAddrMap(Index)) = UCase(RegName) Then
            GetAHBAddress = glGlobalAddrMap(Index)
        End If
    Next Index

    If GetAHBAddress = -1 Then
        TheExec.Datalog.WriteComment "Register Name :   " & RegName & "  =>   is wrong, please take a look!!!"
        '        Stop   '//2019_1213
    End If
    Exit Function

ErrHandler:
    LIB_ErrorDescription ("GetAHBAddress")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


''20180503 evans : Get AHB Field Mask
Public Function GetAHBFieldMask(FieldName As String) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetAHBFieldMask"
    
    Dim Index     As Integer

    GlobalRegMap_Initialize

    GetAHBFieldMask = -1

    For Index = 0 To UBound(gsAHBFieldName)
        If UCase(gsAHBFieldName(Index)) = UCase(FieldName) Then
            GetAHBFieldMask = glAHBFieldMask(Index)
        End If
    Next Index

    Exit Function

ErrHandler:
    LIB_ErrorDescription ("GetAHBFieldMask")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''Public Function GlobalRegMap_Initialize() As Boolean
''
''    On Error GoTo ErrHandler
''    Dim funcName  As String:: funcName = "GlobalRegMap_Initialize"
''
''    Dim lastRowIndex As Range
''    Dim LastRow   As Double
''    Dim RegNameIndex As Integer
''    Dim CheckRegName As String
''    Dim FieldWidth As Long
''    Dim Index     As Integer
''
''    If gbGlobalAddrMap = False Then
''        Dim Row   As Double
''        Dim GlobalAddressMapSheet As Object
''
''        'Find the Last Row Index of the GlobalAddressMap
''        '-------------------------------------------------------------
''        Set GlobalAddressMapSheet = ThisWorkbook.Sheets(gS_AHBRegisterMapSheet)
''        Set lastRowIndex = GlobalAddressMapSheet.Range("A65536").End(xlUp)
''        LastRow = lastRowIndex.Row
''        '-------------------------------------------------------------
''        '20180503 evans.lo : For AHB address
''        RegNameIndex = 0
''        ReDim glGlobalAddrMap(RegNameIndex)
''        ReDim glDictGlobalAddrMap(RegNameIndex)
''        CheckRegName = GlobalAddressMapSheet.Cells(GLOBAL_ADDR_MAP_INDEX.G_START_ROW, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value
''        '20180503 evans.lo : For AHB Field Mask
''        ReDim Preserve gsAHBFieldName(LastRow - GLOBAL_ADDR_MAP_INDEX.G_START_ROW)
''        ReDim Preserve glAHBFieldMask(LastRow - GLOBAL_ADDR_MAP_INDEX.G_START_ROW)
''
''        For Row = GLOBAL_ADDR_MAP_INDEX.G_START_ROW To LastRow
''            '20180503 evans.lo : For AHB address
''            If CheckRegName <> GlobalAddressMapSheet.Cells(Row + 1, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value Then
''                glGlobalAddrMap(RegNameIndex) = CLng(Replace(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_ADDR).Value, "0x", "&H"))
''                glDictGlobalAddrMap(RegNameIndex) = CheckRegName    'GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value
''                CheckRegName = GlobalAddressMapSheet.Cells(Row + 1, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value
''                RegNameIndex = RegNameIndex + 1
''                If Len(CheckRegName) > 0 Then
''                    ReDim Preserve glGlobalAddrMap(RegNameIndex)
''                    ReDim Preserve glDictGlobalAddrMap(RegNameIndex)
''                End If
''            End If
''            '20180503 evans.lo : For AHB Field Mask
''            gsAHBFieldName(Row - GLOBAL_ADDR_MAP_INDEX.G_START_ROW) = UCase(Trim(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_FIELD).Value))
''            FieldWidth = 0
''            For Index = 0 To CLng(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_FIELD_Width)) - 1
''                FieldWidth = FieldWidth + 2 ^ Index
''            Next Index
''            FieldWidth = FieldWidth * 2 ^ CLng(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REF_FIELD_Offset))
''            glAHBFieldMask(Row - GLOBAL_ADDR_MAP_INDEX.G_START_ROW) = CLng("&H" & Mid(CStr(Hex(Not FieldWidth)), 7, 2))
''        Next Row
''        gbGlobalAddrMap = True
''    End If
''
''    GlobalRegMap_Initialize = gbGlobalAddrMap
''
''    Exit Function
''
''ErrHandler:
''    LIB_ErrorDescription ("GlobalRegMap_Initialize")
''    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
''    If AbortTest Then Exit Function Else Resume Next
''
''End Function

'20171027 evans: exort reg data to worksheet or csv file
'20170814 evans: for CT request
Public Function GetDataByRegName_nWire(start_reg As String, end_reg As String, RegCheck As REG_DATA, Optional FileName As String = "REGCHECK.csv", _
                                       Optional ExportToSheet As Boolean = True, Optional ExportToFile As Boolean = True)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetDataByRegName_nWire"


    '''    Dim SiteObject As Object
    '''    Dim ReadExportFileArray() As String
    '''
    '''    On Error GoTo ErrHandler
    '''
    '''    If gbRegDumpOfflineCheck = True Then
    '''        Set SiteObject = TheExec.sites.Existing
    '''    Else
    '''        Set SiteObject = TheExec.sites.Selected
    '''    End If
    '''
    '''    If ExportToSheet = False And ExportToFile = False Then
    '''        TheExec.Datalog.WriteComment "<WARNING> : The alternatives are ExportToSheet and ExportToFile!"
    '''        Exit Function
    '''    End If
    '''
    '''    GlobalRegMap_Initialize
    '''
    '''    If ExportToFile Then
    '''        Dim mS_HEADER As String
    '''        Dim mS_FILETYPE As String
    '''        Dim mS_File  As String
    '''        Dim mS_FileName As String
    '''        Dim fs As New FileSystemObject
    '''        Dim St_ReadTxtFile As TextStream
    '''        Dim read_count As Long
    '''
    '''        mS_FILETYPE = ".csv"
    '''
    '''        mS_FileName = "REGCHECK" + mS_FILETYPE 'Join(mS_FileNameArray, "_") + mS_FILETYPE
    '''        If InStr(mS_FileName, FileName) = 0 And Len(FileName) > 0 Then
    '''            mS_FileName = FileName
    '''            If InStr(FileName, mS_FILETYPE) = 0 Then
    '''                mS_FileName = mS_FileName + mS_FILETYPE
    '''            End If
    '''        End If
    '''
    '''        mS_File = gS_REGCHECKFileDir + mS_FileName
    '''        mS_HEADER = "Site,Reg Name,Reg Addr,Before,After"
    '''
    '''        Call File_CheckAndCreateFolder(gS_REGCHECKFileDir)
    '''        If RegCheck = REG_DATA_BEFORE And fs.FileExists(mS_File) Then
    '''            Call File_CreateAFile(mS_File, mS_HEADER)
    '''        End If
    '''
    '''        read_count = 0
    '''
    '''        If RegCheck = REG_DATA_AFTER Then
    '''
    '''            ReDim FilegReadBySite(TheExec.sites.Existing.Count - 1) 'offline check
    '''
    '''            Set St_ReadTxtFile = fs.OpenTextFile(mS_File, ForReading, True)
    '''            St_ReadTxtFile.ReadLine 'filter header
    '''            Do While Not St_ReadTxtFile.AtEndOfStream
    '''                ReDim Preserve ReadExportFileArray(read_count)
    '''                ReadExportFileArray(read_count) = St_ReadTxtFile.ReadLine
    '''                read_count = read_count + 1
    '''            Loop
    '''            St_ReadTxtFile.Close
    '''            Call ReSortReadRegData(ReadExportFileArray, FilegReadBySite)
    '''        End If
    '''    End If
    '''
    '''    If gbGlobalAddrMap = True Then
    '''        Dim Index As Long
    '''        Dim RegData As New SiteLong
    '''        Dim iSite As Variant
    '''        Dim RegStatusCheckSheet As Object
    '''        Dim Count As Long
    '''        Dim start_index, end_index As Long
    '''        Dim start_addr As Long, end_addr As Long
    '''
    '''        start_addr = -1
    '''        end_addr = -1
    '''
    '''        For Index = 0 To UBound(glDictGlobalAddrMap)
    '''            If glDictGlobalAddrMap(Index) = start_reg Then start_addr = glGlobalAddrMap(Index)
    '''            If glDictGlobalAddrMap(Index) = end_reg Then end_addr = glGlobalAddrMap(Index)
    '''        Next Index
    '''
    '''        If start_addr = -1 Or end_addr = -1 Then
    '''            TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Register Name(it should be from GlobalAddressMap worksheet!)"
    '''            Exit Function
    '''        End If
    '''
    '''        If end_addr < start_addr Then
    '''            TheExec.Datalog.WriteComment "<ERROR> : The end register address should larger than start register address"
    '''            Exit Function
    '''        End If
    '''
    '''        start_index = -1
    '''        end_index = -1
    '''
    '''        For Index = 0 To UBound(glGlobalAddrMap)
    '''            If glGlobalAddrMap(Index) = start_addr Then start_index = Index
    '''            If glGlobalAddrMap(Index) = end_addr Then end_index = Index
    '''        Next Index
    '''
    '''        If start_index = -1 Or end_index = -1 Then
    '''            TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Address(it should be from GlobalAddressMap worksheet!)"
    '''            Exit Function
    '''        End If
    '''
    '''        If (end_index - start_index) > &HFF Then
    '''            TheExec.Datalog.WriteComment "<WARNING> : Dump over 100 registers only export to CSV file!"
    '''            ExportToSheet = False
    '''            ExportToFile = True
    '''        End If
    '''
    '''        Worksheets("REG_STATUS_CHECK").Activate
    '''        Set RegStatusCheckSheet = ThisWorkbook.Sheets("REG_STATUS_CHECK")
    '''
    '''        If RegCheck = REG_DATA_BEFORE Then
    '''            If Len(RegStatusCheckSheet.Cells(2, S_REG_BEFORE).Value) > 0 Then
    '''                RegStatusCheckSheet.UsedRange.ClearContents
    '''            End If
    '''        End If
    '''
    '''
    '''        ReDim ExportRegStatusBySite(TheExec.sites.Existing.Count - 1) 'offline check
    '''        Count = 0
    '''
    '''        For Index = 0 To UBound(glGlobalAddrMap)
    '''            If glGlobalAddrMap(Index) >= start_addr And glGlobalAddrMap(Index) <= end_addr Then
    '''                Call AHB_READNWIRE(glGlobalAddrMap(Index), RegData)
    '''                TheHdw.Wait 0.05
    '''    '''                TheExec.Datalog.WriteComment "RegAddr:" & glDictGlobalAddrMap(Index)
    '''                For Each iSite In SiteObject
    '''                    ReDim Preserve ExportRegStatusBySite(iSite).RegName(Count)
    '''                    ExportRegStatusBySite(iSite).RegName(Count) = glDictGlobalAddrMap(Index)
    '''
    '''                    ReDim Preserve ExportRegStatusBySite(iSite).RegAddr(Count)
    '''                    ExportRegStatusBySite(iSite).RegAddr(Count) = "0x" & CStr(Hex(glGlobalAddrMap(Index)))
    '''
    '''                    If RegCheck = REG_DATA_BEFORE Then
    '''                        ReDim Preserve ExportRegStatusBySite(iSite).BefData(Count)
    '''                        ExportRegStatusBySite(iSite).BefData(Count) = CStr(Hex(RegData(iSite)))
    '''
    '''                        If ExportToFile = True Then
    '''                            ReDim Preserve ExportRegStatusBySite(iSite).ExportToFile(Count)
    '''                            ExportRegStatusBySite(iSite).ExportToFile(Count) = ComposeExportString(iSite, ExportRegStatusBySite(iSite), Count)
    '''                        End If
    '''
    '''                    End If
    '''
    '''                    If RegCheck = REG_DATA_AFTER Then
    '''                        ReDim Preserve ExportRegStatusBySite(iSite).AftData(Count)
    '''                        ExportRegStatusBySite(iSite).AftData(Count) = CStr(Hex(RegData(iSite)))
    '''
    '''                        If ExportToFile = True Then
    '''                            If CompareExportString(ReadExportFileArray(Count), glDictGlobalAddrMap(Index)) = False Then
    '''                                TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Address(it should be same as Previous Address Setup!)"
    '''                                Exit Function
    '''                            Else
    '''                                ReDim Preserve ExportRegStatusBySite(iSite).ExportToFile(Count)
    '''                                ExportRegStatusBySite(iSite).ExportToFile(Count) = FilegReadBySite(iSite).READDATA(Count) + ExportRegStatusBySite(iSite).AftData(Count) _
     '''                                                                                  + ComposeRegData(FilegReadBySite(iSite).READDATA(Count), ExportRegStatusBySite(iSite).AftData(Count))
    '''                            End If
    '''                        End If
    '''
    '''                    End If
    '''
    '''                Next iSite
    '''                Count = Count + 1
    '''            End If
    '''        Next Index
    '''
    '''        If ExportToSheet Then
    '''            Index = 0
    '''            Count = 2
    '''            RegStatusCheckSheet.Cells(Index + 1, S_REG_SITE).Value = "Site"
    '''            RegStatusCheckSheet.Cells(Index + 1, S_REG_NAME).Value = "Reg Name"
    '''            RegStatusCheckSheet.Cells(Index + 1, S_REG_ADDR).Value = "Reg Addr"
    '''            RegStatusCheckSheet.Cells(Index + 1, S_REG_BEFORE).Value = "Before"
    '''            RegStatusCheckSheet.Cells(Index + 1, S_REG_AFTER).Value = "After"
    '''            For Each iSite In SiteObject
    '''                For Index = 0 To UBound(ExportRegStatusBySite(iSite).RegName)
    '''                    RegStatusCheckSheet.Cells(Index + Count, S_REG_SITE).Value = CStr(iSite)
    '''                    RegStatusCheckSheet.Cells(Index + Count, S_REG_NAME).Value = ExportRegStatusBySite(iSite).RegName(Index)
    '''                    RegStatusCheckSheet.Cells(Index + Count, S_REG_ADDR).Value = "0x" & ExportRegStatusBySite(iSite).RegAddr(Index)
    '''                    If RegCheck = REG_DATA_BEFORE Then RegStatusCheckSheet.Cells(Index + Count, S_REG_BEFORE).Value = ExportRegStatusBySite(iSite).BefData(Index)
    '''                    If RegCheck = REG_DATA_AFTER Then
    '''                        RegStatusCheckSheet.Cells(Index + Count, S_REG_AFTER).Value = ExportRegStatusBySite(iSite).AftData(Index)
    '''                        If RegStatusCheckSheet.Cells(Index + Count, S_REG_AFTER).Value <> RegStatusCheckSheet.Cells(Index + Count, S_REG_BEFORE).Value Then
    '''                            RegStatusCheckSheet.Cells(Index + Count, S_REG_CHECK).Value = DiffCheck
    '''                        End If
    '''                    End If
    '''                Next Index
    '''                Count = Count + UBound(ExportRegStatusBySite(iSite).RegName) + 1
    '''            Next iSite
    '''        End If
    '''
    '''        If ExportToFile Then
    '''            Dim St_WriteTxtFile As TextStream
    '''
    '''            Set St_WriteTxtFile = fs.OpenTextFile(mS_File, ForWriting, True)
    '''            St_WriteTxtFile.WriteLine mS_HEADER
    '''            For Each iSite In SiteObject
    '''                For Index = 0 To UBound(ExportRegStatusBySite(iSite).ExportToFile)
    '''                   St_WriteTxtFile.WriteLine ExportRegStatusBySite(iSite).ExportToFile(Index)
    '''                Next Index
    '''            Next iSite
    '''
    '''            St_WriteTxtFile.Close
    '''
    '''        End If
    '''
    '''    End If
    '''
    '''    TheExec.Datalog.WriteComment "<GetDataByRegName> : Register dump is completed!"
    '''
    '''    Exit Function

ErrHandler:
    LIB_ErrorDescription ("GetDataByRegName")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function




Private Function CompareExportString(ReadExportString As String, RegName As String) As Boolean

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "CompareExportString"
    
    If InStr(ReadExportString, RegName) <> 0 Then
        CompareExportString = True
    Else
        CompareExportString = False
    End If

    Exit Function
ErrHandler:
    LIB_ErrorDescription ("CompareExportString")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function




'20190122
'To compare the digital pins connecting status
'call the function Check_DigCh_ConnState CHECK_DATA_Before as the baseline
'and the results will be printed while Check_DigCh_ConnState CHECK_DATA_After is executed.
Public Function Check_DigCh_ConnState(Optional CHECKSTAGE As CHECK_DATA, Optional ExportToFile As Boolean = True)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Check_DigCh_State"
    Dim mS_Title_Before As String
    Dim mS_Title_After As String
    Dim mS_Title  As String
    Dim mS_Temp   As String
    Dim PinName() As String
    Dim PinCnt    As Long
    Dim idx       As Long
    Dim channels  As String
    Dim siteCnt   As Double
    Dim DigitalDiffCount As Integer
    mS_Title_Before = ""
    mS_Title_After = ""
    TheExec.DataManager.DecomposePinList "ALL_DIG_PINS_NO_FRC", PinName(), PinCnt
    TheExec.DataManager.ReturnSignalNames = True    'use signal style

    siteCnt = TheExec.Sites.Selected.Count

    For Each g_Site In TheExec.Sites.Selected
        mS_Title_Before = mS_Title_Before & vbTab & "GetStatus_Before(S" & g_Site & ")"
        mS_Title_After = mS_Title_After & vbTab & "GetStatus_After(S" & g_Site & ")"
    Next g_Site
    mS_Title = FormatLog("PinName", -20) & "," & mS_Title_Before & _
               "," & mS_Title_After

    'Debug.Print mS_Title

    Dim ConnState() As Boolean
    ReDim ConnState(PinCnt - 1, siteCnt)

    If CHECKSTAGE = CHECK_DATA_BEFORE Then
        ReDim g_DigiState_Stru.State_Before(PinCnt - 1)
        '___Initalize
        For idx = 0 To PinCnt - 1
            g_DigiState_Stru.State_Before(idx) = ""
        Next idx
    Else
        ReDim g_DigiState_Stru.State_After(PinCnt - 1)
    End If

    For idx = 0 To PinCnt - 1
        For Each g_Site In TheExec.Sites
            'theexec.DataManager.GetChannelStringFromPinAndSite pinName(idx), g_Site, channels
            'ConnState = thehdw.Digital.Raw.Chans(channels).isConnected
            ConnState(idx, g_Site) = TheHdw.Digital.Raw.Chans(PinName(idx)).IsConnected
            If CHECKSTAGE = CHECK_DATA_BEFORE Then
                g_DigiState_Stru.State_Before(idx) = g_DigiState_Stru.State_Before(idx) & FormatLog(ConnState(idx, g_Site), -25)
            ElseIf CHECKSTAGE = CHECK_DATA_After Then
                If idx = 0 Then DigitalDiffCount = 0
                g_DigiState_Stru.State_After(idx) = g_DigiState_Stru.State_After(idx) & FormatLog(ConnState(idx, g_Site), -25)
            End If
        Next

        'Datalog:
        If CHECKSTAGE = CHECK_DATA_BEFORE Then
            'mS_Temp = FormatLog(pinName(idx), -20) & "," & FormatLog(g_DigiState_Stru.State_Before(idx), -25)
            'TheExec.Datalog.WriteComment mS_Temp
        ElseIf CHECKSTAGE = CHECK_DATA_After Then
            If idx = 0 Then TheExec.Datalog.WriteComment mS_Title
            If g_DigiState_Stru.State_After(idx) <> g_DigiState_Stru.State_Before(idx) Then
                mS_Temp = FormatLog(PinName(idx), -20) & "," & FormatLog(g_DigiState_Stru.State_Before(idx), -25) & _
                          "," & FormatLog(g_DigiState_Stru.State_After(idx), -25)
                TheExec.Datalog.WriteComment mS_Temp
                DigitalDiffCount = DigitalDiffCount + 1
            End If
            If idx = PinCnt - 1 Then
                If DigitalDiffCount = 0 Then
                    TheExec.Datalog.WriteComment "Digital connection status are identical !!!"
                Else
                    TheExec.Datalog.WriteComment "There are " & DigitalDiffCount & "different digital connection status !!!"
                End If
            End If
        End If
    Next idx


    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function



