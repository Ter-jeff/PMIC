Attribute VB_Name = "LIB_Common"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit
'Revision History:
'V0.0 initial bring up
'V0.1 add bintable inital VBT.
'variable declaration
Public Const Version_Lib_Common As String = "0.1" 'lib version
Public Const DebugPrintEnable As Boolean = False  'debug print in VBT modules
Public G_TestName As String 'replace testinstance for debug print CHWu 102615
Public Current_Patterns As String
Public Char_Test_Name_Curr_Loc As Long 'index for char datalog test name array
Private Type Bintable
    astrBinName() As String
    astrBinRename() As String
    astrBinSortNum() As String
End Type
Public m1_InstanceName As String

Public tyBinTable As Bintable

Public nWire_Ports_GLB As String ''Support multiple nWire port 20170718

Public Previous_DCCategory As String
Public Previous_DCSelector As String

'20170814 evans.lo: for CT request
Public gbGlobalAddrMap As Boolean
Public glGlobalAddrMap() As Long
Public glDictGlobalAddrMap() As String
'20180503 evans.lo: for Avus field mask request
Dim gsAHBFieldName() As String
Dim glAHBFieldMask() As Long

'20171027 evans: for Reg debug check
Const gbRegDumpOfflineCheck As Boolean = False 'It's for offline check

Private Enum GLOBAL_ADDR_MAP_INDEX
    G_START_ROW = 2
    G_REG_NAME = 5
    G_REG_ADDR = 3
    G_REG_FIELD = 6
    G_REG_FIELD_Width = 7
    G_REF_FIELD_Offset = 8
End Enum

'Public Enum REG_DATA
'    REG_DATA_BEFORE
'    REG_DATA_AFTER
'End Enum

Private Const gS_REGCHECKFileDir = ".\REGCHECK\"

Private Type REG_FILE_READ
    READDATA() As String
End Type

Private FilegReadBySite() As REG_FILE_READ

Private Enum REG_STATUS_INDEX
    S_REG_SITE = 1
    S_REG_NAME = 2
    S_REG_ADDR = 3
    S_REG_BEFORE = 4
    S_REG_AFTER = 5
    S_REG_CHECK = 6
End Enum

Public Enum REG_DATA
    REG_DATA_BEFORE
    REG_DATA_AFTER
End Enum

Private Type REG_STATUS_EXPORT
    RegName() As String
    RegAddr() As String
    BefData() As String
    AftData() As String
    ExportToFile() As String
End Type

Private ExportRegStatusBySite() As REG_STATUS_EXPORT
Private Const DiffCheck = "Y"

'20170808 evans.lo

Public Const gDCTestNameTemplate As String = "X_X_X_X_X_X_X_X_X_X_X_X"
Public Enum DC_TNAME_INDEX
    DC_TNAME_TESTITEM = 0
    DC_TNAME_TESTPIN = 1
    DC_TNAME_TESTCOND = 4
    DC_TNAME_TESTTEMP = 11
End Enum
'Private gDCTestNameIndex As Long
Public gDCArrTestName() As String
Public TPModeAsCharz_GLB As Boolean 'VBT_LIB_Digital_Shmoo
Public Enum dcvi_type
    DCVI_DC30 = 1
    DCVI_UVI80 = 2
    DCVI_DC75 = 3     ' CURRENTLY NOT IN CONFIG
End Enum

Public Enum YESNO_type
    Yes = 0
    No = 1
End Enum



'*****************************************
'******                ascii utility******
'*****************************************
Public Sub AsciiUtilsPreBatchFileCreation()
    'asciiutils.Control.OverrideRelativeSrcPath "./src"
End Sub

Public Sub AsciiUtilsPreSaveToASCII()
   ' asciiutils.Tools.SetCodeCase
    'asciiutils.Control.OverrideRelativeSrcPath "./src"
End Sub
Function Is64bit() As Boolean
    Is64bit = Len(Environ("ProgramW6432")) > 0
End Function
Function DoesReferenceExist(ByVal RName As String) As Boolean
    On Error Resume Next
    Dim Ref As Variant
    DoesReferenceExist = False
    For Each Ref In ThisWorkbook.VBProject.References
        'Debug.Print ref.Name
        If Ref.Name = RName Then
            DoesReferenceExist = True
        End If
    Next

End Function
'Public Function AddRef()
'On Error GoTo ErrHandler
'Dim funcName As String:: funcName = "AddRef"
'
'    If Is64bit Then
'        If DoesReferenceExist("Scripting") Then
'            Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\WINDOWS\sysWOW64\scrrun.dll"
'        End If
'        If DoesReferenceExist("PATTERNDATAMANAGERLib") Then
'            Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\Program Files (x86)\Teradyne\IG-XL\8.10.12_uflx\bin\PatternDataManager.dll"
'        End If
'    Else
'        If DoesReferenceExist("Scripting") Then
'            Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\WINDOWS\system32\scrrun.dll"
'        End If
'        If DoesReferenceExist("PATTERNDATAMANAGERLib") Then
'            Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\Program Files\Teradyne\IG-XL\8.10.90_uflx\bin\PatternDataManager.dll"
'        End If
'    End If
'
'Exit Function
'
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function
Public Function is_reference_installed(s As String) As Boolean

On Error GoTo ErrHandler
Dim funcName As String:: funcName = "is_reference_installed"

    Dim x As Variant
    is_reference_installed = False
    For Each x In Application.ActiveWorkbook.VBProject.References
        If s = x.Name Then
            is_reference_installed = True
        End If
    Next x
Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Function CheckAddin(s As String) As Boolean
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "CheckAddin"

    Dim x As Variant
    On Error Resume Next
    x = AddIns(s).Installed
    On Error GoTo 0
    If IsEmpty(x) Then
        CheckAddin = False
    Else
        CheckAddin = True
    End If
    
Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function
Function Print_Addin() As Boolean
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Print_Addin"

    Dim x As Variant
    Dim oAddIn As AddIn
    
    Debug.Print "------ Add In"
    For Each oAddIn In Application.AddIns
        Debug.Print oAddIn.Name
    Next oAddIn
       
    Debug.Print "------ Reference"
    For Each x In Application.ActiveWorkbook.VBProject.References
        Debug.Print x.Name
    Next x
Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Function WorksheetExists(wsName As String, delete As Boolean) As Boolean
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "WorksheetExists"

    Dim ws As Worksheet
    Dim ret As Boolean
    ret = False
    wsName = UCase(wsName)
    For Each ws In ThisWorkbook.Sheets
        If UCase(ws.Name) = wsName Then
            If delete = True Then
                Application.DisplayAlerts = False
                ws.delete
                Application.DisplayAlerts = True
            End If
            ret = True
            Exit For
        End If
    Next
    WorksheetExists = ret

Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Wait(Time As Double, Optional Debug_Flag As Boolean = False)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Wait"
'pause few time
    TheHdw.Wait Time
    If Debug_Flag Then
        TheExec.Datalog.WriteComment ("print: Wait time = " + CStr(Time * 1000#) + " mS")
    End If
    
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub ErrorDescription(funcName As String)
On Error GoTo ErrHandler
'error description printing
    Dim TestInstanceName As String
    TestInstanceName = TheExec.DataManager.InstanceName
    
    TheExec.Datalog.WriteComment "TestInstance: " & TestInstanceName & ", " & funcName & " error, Err Code: " & err.Number & ", Err Description: " & err.Description

Exit Sub
ErrHandler:
     If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function Trim_NC_Pin(ByRef original_ary() As String, ByRef original_pin_cnt As Long)
'get active pins array
    Dim i As Long, j As Long
    Dim p As Variant
    Dim TempArray() As String
    Dim TempPinCnt As Long
    Dim NullArray() As String
    Dim TempString As String
    Dim PowerSequence As Double
    
    On Error GoTo ErrHandler
    
    If original_pin_cnt <> 0 Then
        i = 0   'init
        For Each p In original_ary
            If TheExec.DataManager.ChannelType(p) <> "N/C" Then i = i + 1
        Next p
        
        'redim
        ReDim TempArray(i - 1)
        
        j = 0   'init
        
        If i > 0 Then
            'redim
            ReDim TempArray(i - 1)
            
            For Each p In original_ary
                If TheExec.DataManager.ChannelType(p) <> "N/C" Then
                    TempArray(j) = p
                    j = j + 1
                Else
                    j = j
                End If
            Next p
        End If

        For Each p In original_ary
            If TheExec.DataManager.ChannelType(p) <> "N/C" Then
                TempArray(j) = original_ary(j)
                j = j + 1
            Else
                j = j
            End If
        Next p
        
        'return array and pin count
        original_ary = TempArray
        original_pin_cnt = j
    End If
    
    Exit Function
ErrHandler:
    ErrorDescription ("Trim_NC_Pin")
    If AbortTest Then Exit Function Else Resume Next
End Function

'Public Sub RemoveandCopyQualifiers()
'On Error GoTo ErrHandler
'Dim funcName As String:: funcName = "RemoveandCopyQualifiers"
'
'
'    Dim n As Long
'    Dim i As Long
'    Dim Header As String
'    Dim WorkSheetType As String
'    Dim Posn As Long
'    Dim Row As Long
'    Dim Opcode As String
'    Dim Continue As Boolean
'    Dim usedrow As Double
'
'    n = Worksheets.Count
'
'    For i = 1 To n
'    Header = Worksheets(i).Cells(1, 1).Value
'
'    Posn = InStr(Header, ",")
'    If Posn > 0 Then
'    WorkSheetType = Mid(Header, 1, Posn - 1)
'    If WorkSheetType = "DTFlowtableSheet" Then
'    Worksheets(i).Activate
'    usedrow = ActiveSheet.UsedRange.Rows.Count
'    ActiveSheet.Range(Cells(5, 24), Cells(usedrow, 30)).Select
'    ActiveSheet.Range(Cells(5, 24), Cells(usedrow, 30)).Copy
'    ActiveSheet.Range(Cells(5, 44), Cells(usedrow, 50)).Select
'    ActiveSheet.Range(Cells(5, 44), Cells(usedrow, 50)).PasteSpecial
'    ActiveSheet.Range(Cells(5, 24), Cells(usedrow, 30)).Select
'    ActiveSheet.Range(Cells(5, 24), Cells(usedrow, 30)).ClearContents
'    End If
'    End If
'
'    Next i
'
'Exit Sub
'ErrHandler:
'     RunTimeError funcName
'     If AbortTest Then Exit Sub Else Resume Next
'End Sub
'
'Public Sub RecoveryandCopyQualifiers()
'On Error GoTo ErrHandler
'Dim funcName As String:: funcName = "RecoveryandCopyQualifiers"
'
'    Dim n As Long
'    Dim i As Long
'    Dim Header As String
'    Dim WorkSheetType As String
'    Dim Posn As Long
'    Dim Row As Long
'    Dim Opcode As String
'    Dim Continue As Boolean
'    Dim usedrow As Double
'
'    n = Worksheets.Count
'
'    For i = 1 To n
'    Header = Worksheets(i).Cells(1, 1).Value
'
'    Posn = InStr(Header, ",")
'    If Posn > 0 Then
'    WorkSheetType = Mid(Header, 1, Posn - 1)
'    If WorkSheetType = "DTFlowtableSheet" Then
'    Worksheets(i).Activate
'    usedrow = ActiveSheet.UsedRange.Rows.Count
'    ActiveSheet.Range(Cells(5, 44), Cells(usedrow, 50)).Select
'    ActiveSheet.Range(Cells(5, 44), Cells(usedrow, 50)).Copy
'    ActiveSheet.Range(Cells(5, 24), Cells(usedrow, 30)).Select
'    ActiveSheet.Range(Cells(5, 24), Cells(usedrow, 30)).PasteSpecial
'    ActiveSheet.Range(Cells(5, 44), Cells(usedrow, 50)).Select
'    ActiveSheet.Range(Cells(5, 44), Cells(usedrow, 50)).ClearContents
'    End If
'    End If
'
'    Next i
'
'Exit Sub
'ErrHandler:
'     RunTimeError funcName
'     If AbortTest Then Exit Sub Else Resume Next
'End Sub
'
''*****************************************
''******                         TDR ******
''*****************************************
'Public Function TDR_Gen_Cal_File() As Long
'On Error GoTo ErrHandler
'Dim funcName As String:: funcName = "TDR_Gen_Cal_File"
'
'    Dim ws_chan As Worksheet
'    Dim ws_tdr As Worksheet
'    Dim wb As Workbook
'    Dim chanmap_name As String, tdr_name As String
'    Dim row_chan As Long, row_tdr As Long
'    Dim Site As Variant
'
'    TheExec.Datalog.WriteComment "<" & TheExec.DataManager.InstanceName & ">"
'
'    Set wb = Application.ActiveWorkbook
'    chanmap_name = TheExec.CurrentChanMap
'    Set ws_chan = wb.Sheets(chanmap_name)
'    ws_chan.Cells(3, 6) = "Signal"
'    tdr_name = "TDR_DATA_" & chanmap_name
'    Call WorksheetExists(tdr_name, True)
'    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = tdr_name
'     Set ws_tdr = wb.Sheets(tdr_name)
'
'    ws_tdr.Cells(1, 1).Value = "Pin Name"
'    ws_tdr.Cells(1, 1).Interior.Color = RGB(128, 128, 0)
'    For Each Site In TheExec.Sites.Existing
'        ws_tdr.Cells(1, 2).Value = "Chan Site" & Site
'        ws_tdr.Cells(1, 2).Interior.Color = RGB(128, 128, 0)
'        ws_tdr.Cells(1, CLng(Site) + 3).Value = "Trace Site " & Site
'        ws_tdr.Cells(1, CLng(Site) + 3).Interior.Color = RGB(128, 128, 0)
'    Next Site
'
'    row_chan = 7
'    row_tdr = 2
'    While (ws_chan.Cells(row_chan, 2).Value <> "")
'        If ws_chan.Cells(row_chan, 4).Value = "I/O" Then
'            For Each Site In TheExec.Sites.Existing
'                    ws_tdr.Cells(row_tdr, 1).Value = ws_chan.Cells(row_chan, 2).Value
'                    If ws_chan.Cells(row_chan, 5 + CLng(Site)).Value Like "*.ch*" Then
'                        ws_tdr.Cells(row_tdr, 2 + 2 * CLng(Site)).Value = ws_chan.Cells(row_chan, 5 + CLng(Site)).Value
'                        ws_tdr.Cells(row_tdr, 3 + 2 * CLng(Site)).Value = TheHdw.Digital.Calibration.channels(ws_chan.Cells(row_chan, 5 + CLng(Site)).Value).DIB.Trace
'                    ElseIf ws_chan.Cells(row_chan, 5 + CLng(Site)).Value Like "*site*" Then
'                        ws_tdr.Cells(row_tdr, 2 + 2 * CLng(Site)).Value = ws_chan.Cells(row_chan, 5 + CLng(0)).Value
'                        ws_tdr.Cells(row_tdr, 3 + 2 * CLng(Site)).Value = TheHdw.Digital.Calibration.channels(ws_chan.Cells(row_chan, 5 + CLng(0)).Value).DIB.Trace
'                    Else
'                        MsgBox " select Signal in chanmap sheet  " & chanmap_name
'                        Exit Function
'                    End If
'                    If ws_tdr.Cells(row_tdr, 3 + 2 * CLng(Site)).Value > 0.0000000655 Then
'                        ws_tdr.Cells(row_tdr, 3 + 2 * CLng(Site)).Interior.Color = RGB(255, 0, 0)
'                    Else
'                        ws_tdr.Cells(row_tdr, 3 + 2 * CLng(Site)).Interior.Color = RGB(255, 255, 255)
'                    End If
'
'            Next Site
'            row_tdr = row_tdr + 1
'        End If
'        row_chan = row_chan + 1
'    Wend
'
'    Exit Function
'ErrHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'Public Function TDR_Read_Compare() As Long
'    On Error GoTo ErrHandler
'    Dim ws_chan As Worksheet
'    Dim ws_tdr As Worksheet, ws_tdr_cmp As Worksheet
'    Dim wb As Workbook
'    Dim chanmap_name As String, tdr_name As String, tdr_cmp_name As String
'    Dim row_chan As Long, row_tdr As Long, tdr_chan As String
'    Dim Site As Variant
'
'    TheExec.Datalog.WriteComment "<" & TheExec.DataManager.InstanceName & ">"
'
'    Set wb = Application.ActiveWorkbook
'    chanmap_name = TheExec.CurrentChanMap
'    Set ws_chan = wb.Sheets(chanmap_name)
'
'    tdr_name = "TDR_DATA_" & chanmap_name
'    If WorksheetExists(tdr_name, False) = False Then
'        TheExec.AddOutput tdr_name & " does not exist"
'        Exit Function
'    Else
'        Set ws_tdr = wb.Sheets(tdr_name)
'    End If
'
'    tdr_cmp_name = "TDR_CMPARE"
'    WorksheetExists tdr_cmp_name, True
'    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = tdr_cmp_name
'    Set ws_tdr_cmp = wb.Sheets(tdr_cmp_name)
'
'    ws_tdr_cmp.Cells(1, 1).Value = "Pin Name"
'    ws_tdr_cmp.Cells(1, 1).Interior.Color = RGB(128, 128, 0)
'
'    For Each Site In TheExec.Sites.Existing
'        ws_tdr_cmp.Cells(1, 2).Value = "Chan Site" & Site
'        ws_tdr_cmp.Cells(1, 2).Interior.Color = RGB(128, 128, 0)
'        ws_tdr_cmp.Cells(1, CLng(Site) + 3).Value = "Org Trace Site " & Site
'        ws_tdr_cmp.Cells(1, CLng(Site) + 3).Interior.Color = RGB(128, 128, 0)
'        ws_tdr_cmp.Cells(1, CLng(Site) + 4).Value = "Tester Trace Site " & Site
'        ws_tdr_cmp.Cells(1, CLng(Site) + 4).Interior.Color = RGB(128, 128, 0)
'    Next Site
'
'    row_tdr = 2
'    While (ws_tdr.Cells(row_tdr, 2).Value <> "")
'        ws_tdr_cmp.Cells(row_tdr, 1).Value = ws_tdr.Cells(row_tdr, 1).Value
'        For Each Site In TheExec.Sites.Existing
'                tdr_chan = ws_tdr.Cells(row_tdr, 2 + 2 * CLng(Site)).Value
'                ws_tdr_cmp.Cells(row_tdr, 2 + 3 * CLng(Site)).Value = tdr_chan
'                ws_tdr_cmp.Cells(row_tdr, 3 + 3 * CLng(Site)).Value = ws_tdr.Cells(row_tdr, 3 + 2 * CLng(Site)).Value
'                ws_tdr_cmp.Cells(row_tdr, 4 + 3 * CLng(Site)).Value = TheHdw.Digital.Calibration.channels(tdr_chan).DIB.Trace
'                If Abs(ws_tdr_cmp.Cells(row_tdr, 3 + 3 * CLng(Site)).Value - ws_tdr_cmp.Cells(row_tdr, 4 + 3 * CLng(Site)).Value) > 0.0000000001 Then
'                    ws_tdr_cmp.Cells(row_tdr, 4 + 3 * CLng(Site)).Interior.Color = RGB(255, 0, 0)
'                Else
'                    ws_tdr_cmp.Cells(row_tdr, 4 + 3 * CLng(Site)).Interior.Color = RGB(255, 255, 255)
'                End If
'        Next Site
'        row_tdr = row_tdr + 1
'    Wend
'
'    Exit Function
'ErrHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'Public Function TDR_Write() As Long
'    Dim ws_chan As Worksheet
'    Dim ws_tdr As Worksheet, ws_tdr_cmp As Worksheet
'    Dim wb As Workbook
'    Dim chanmap_name As String, tdr_name As String, tdr_cmp_name As String
'    Dim row_chan As Long, row_tdr As Long, tdr_chan As String
'    Dim Site As Variant
'    Dim tdr_len As Double, tdr_pin As String, row_change_start As Double, row_change_stop As Double
'    On Error GoTo ErrHandler
'
'    TheExec.Datalog.WriteComment "<" & TheExec.DataManager.InstanceName & ">"
'    Set wb = Application.ActiveWorkbook
'    chanmap_name = TheExec.CurrentChanMap
'    Set ws_chan = wb.Sheets(chanmap_name)
'
'    tdr_name = "TDR_DATA_" & chanmap_name
'    If WorksheetExists(tdr_name, False) = False Then
'        TheExec.AddOutput tdr_name & " does not exist"
'        Exit Function
'    Else
'        Set ws_tdr = wb.Sheets(tdr_name)
'    End If
'
'    row_tdr = 2
'     If LCase(TheExec.CurrentJob) Like "ft*" Then
'        row_change_start = 163 - 6 + 1     'refer to  chanmap
'        row_change_stop = 188 - 6 + 1      'refer to  chanmap
'    ElseIf LCase(TheExec.CurrentJob) Like "cp*" Then
'        row_change_start = 163 - 6 + 1      'refer to  chanmap
'        row_change_stop = 188 - 5 - 6 + 1   'refer to  chanmap
'    End If
'
'    While (ws_tdr.Cells(row_tdr, 2).Value <> "")
'        For Each Site In TheExec.Sites.Existing
'                tdr_chan = ws_tdr.Cells(row_tdr, 2 + 2 * CLng(Site)).Value
'                tdr_len = ws_tdr.Cells(row_tdr, 3 + 2 * CLng(Site)).Value
'                If row_tdr < row_change_start Or row_tdr > row_change_stop Then
''                    theexec.AddOutput "site " & site & "," & row_tdr - 1 & "," & ws_tdr.Cells(row_tdr, 1).Value & "," & _
''                                                            tdr_chan & "," & _
''                                                            ws_tdr.Cells(row_tdr, 3 + 2 * CLng(site)).Value
'                    TheHdw.Digital.Calibration.channels(tdr_chan).DIB.Trace = tdr_len
'                Else 'overwrite MIPI LP with MIPI HS
''                     theexec.AddOutput "site " & site & "," & row_tdr - 1 & "," & ws_tdr.Cells(row_tdr - 26, 1).Value & "," & _
''                                                            tdr_chan & "," & _
''                                                            ws_tdr.Cells(row_tdr - 26, 3 + 2 * CLng(site)).Value
'                    TheHdw.Digital.Calibration.channels(tdr_chan).DIB.Trace = ws_tdr.Cells(row_tdr - 26, 3 + 2 * CLng(Site)).Value
'               End If
'        Next Site
'        row_tdr = row_tdr + 1
'    Wend
'    TheExec.Datalog.WriteComment "num of pins = " & row_tdr - 2
'
'
'    Exit Function
'ErrHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function ReadWaferData()
'
'    Dim LotTmp As String, WaferTmp As String
'    Dim X_Tmp As String, Y_Tmp As String
'    Dim Loc_dash As Integer
'    Dim Site As Variant
'
'    On Error GoTo err1
'
'    '=== Initialization of parameters ====
'
'    '=== Simulated Data ===
'    If (TheExec.TesterMode = testModeOffline) Then
'        LotTmp = "N99G19-01E0"
'    Else
'        LotTmp = TheExec.Datalog.Setup.LotSetup.LotID
'
'    End If
'
'    Loc_dash = InStr(1, LotTmp, "-")
'
'    If Loc_dash <> 0 Then
'        LotID = Mid(LotTmp, 1, Loc_dash - 1)
'    Else
'        LotID = LotTmp
'    End If
'
'    If (TheExec.TesterMode = testModeOffline) Then
'        WaferID = Mid(LotTmp, Loc_dash + 1, 2)
'    Else
'        If TheExec.Datalog.Setup.WaferSetup.ID <> "" Then WaferID = TheExec.Datalog.Setup.WaferSetup.ID
'    End If
'
'    For Each Site In TheExec.Sites
'        If (TheExec.TesterMode = testModeOffline) Then
'            XCoord(Site) = 1
'            YCoord(Site) = 11 - Site
'        Else
'            XCoord(Site) = TheExec.Datalog.Setup.WaferSetup.GetXCoord(Site)
'            YCoord(Site) = TheExec.Datalog.Setup.WaferSetup.GetYCoord(Site)
'
'            If (XCoord(Site) = -32768) Then
'                XCoord(Site) = 1
'                YCoord(Site) = 11 - Site
'            End If
'        End If
'
'        TheExec.Datalog.WriteComment "Lot ID = " + LotID
'        TheExec.Datalog.WriteComment "Wafer ID = " + CStr(WaferID)
'        TheExec.Datalog.WriteComment "X coor (site " + CStr(Site) + ")= " + CStr(XCoord(Site))
'        TheExec.Datalog.WriteComment "Y coor (site " + CStr(Site) + ")= " + CStr(YCoord(Site))
'        'TheExec.Datalog.WriteComment "DFT Type = " & g_sDFT_Type(Site)
'    Next Site
'    Exit Function
'err1:
'    TheExec.Datalog.WriteComment ("There is an error happened in the function of ReadWaferData()")
'                If AbortTest Then Exit Function Else Resume Next
'End Function
'Public Function ReadHandlerData()
'On Error GoTo ErrHandler
'Dim funcName As String:: funcName = "ReadHandlerData"
'
''update by Jason's request to fixed Galaxy multi-site format, 140411
'    Dim Site As Variant
'    For Each Site In TheExec.Sites
'        TheExec.Datalog.WriteComment ("<@Chuck_ID=" & Site & "|" & RegKeyRead("Cover_ID") & ">")
'        TheExec.Datalog.WriteComment ("<@Dut_Temperature=" & Site & "|" & RegKeyRead("Dut_Temperature") & ">")
'        TheExec.Datalog.WriteComment ("<@Handler_Arm_ID=" & Site & "|" & RegKeyRead("Handler_Arm_ID") & ">")
'        TheExec.Datalog.WriteComment ("<@Rework_Flag=" & Site & "|" & RegKeyRead("FT_ReTest") & ">")
'        TheExec.Datalog.WriteComment ("<@Socket_ID=" & Site & "|" & RegKeyRead("Socket_ID") & ">")
'        TheExec.Datalog.WriteComment ("<@Sort_Stage=" & Site & "|" & RegKeyRead("Sort_Code") & ">")
'    Next Site
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'Public Function ShowECIDData()
'On Error GoTo ErrHandler
'Dim funcName As String:: funcName = "ShowECIDData"
''update by Jason's request to fixed Galaxy multi-site format, 140411
'    Dim Site As Variant
'    For Each Site In TheExec.Sites
'        TheExec.Datalog.WriteComment "<@efuse_lot_ID=" & Site & "|" & HramLotId(Site) & ">"
'        TheExec.Datalog.WriteComment "<@efuse_wafer_ID=" & Site & "|" & HramLotId(Site) & "." & Format(CStr(HramWaferId(Site)), "00") & ">"
'    Next Site
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function
Public Function RegKeySave(i_RegKey As String, i_Value As String, Optional i_Type As String = "REG_SZ")
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "RegKeySave"

'add by Teradyne/Vern to output variable to RegKey
'sets the registry key i_RegKey to the
'value i_Value with type i_Type
'if i_Type is omitted, the value will be saved as string
'if i_RegKey wasn't found, a new registry key will be created
        Dim myWS As Object
        'access Windows scripting
        Set myWS = CreateObject("WScript.Shell")
        'write registry key
        i_RegKey = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\" & i_RegKey
        myWS.RegWrite i_RegKey, i_Value, i_Type
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Function RegKeyRead(i_RegKey As String) As String

Dim myWS As Object

On Error Resume Next

Set myWS = CreateObject("WScript.Shell")

RegKeyRead = myWS.RegRead("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\" & i_RegKey)

End Function
Public Function Dec2Bin(ByVal n As Long, ByRef BinArray() As Long)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Dec2Bin"

    Dim i As Integer, j As Integer
    Dim Element_Amount As Integer
    Dim Count As Integer
    '               01101
    ' BinArray(4) 1
    ' BinArray(3) 0
    ' BinArray(2) 1
    ' BinArray(1) 1
    ' BinArray(0) 0

    Element_Amount = UBound(BinArray)
    If n > (2 ^ (Element_Amount + 1) - 1) Then
        n = 0
        TheExec.Datalog.WriteComment "Error(Dec2Bin): Overange for " & n
    End If

    For j = 0 To Element_Amount
        BinArray(j) = 0
    Next j

    'If n < 0 Then MsgBox ("Warning(Dec2Bin)!!! Decimal Number should be positive integer")
    If n < 0 Then
        TheExec.Datalog.WriteComment " The input vlaue of (Dec2Bin) is negative, so we enforce it as 0 to prevent from error alarm."
        n = 0
    End If
    
    i = 0
    Do Until n = 0
        If (i > Element_Amount) Then TheExec.Datalog.WriteComment "Warning (Dec2Bin)!!! Decimal " & n & " is over-range (>" & i & "bit)"
        If (n Mod 2) Then
            BinArray(Element_Amount - i) = 1
        Else
            BinArray(Element_Amount - i) = 0
        End If
        n = Int(n / 2)
        i = i + 1
    Loop

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Dec2BinStr32Bit(ByVal Nbit As Long, ByVal num As Long) As String
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Dec2BinStr32Bit"
    ' 2'complement: invert the number's bits and then add 1
    'Dec2BinStr32Bit 32, -65525
    '1111111111111110000000000001011    -65525
    '0000000000000001111111111110101     65525
    Dim i As Integer, j As Integer
    Dim Element_Amount As Integer
    Dim Count As Integer
    Dim BinStr As String
    ' MSB "010101" LSB
    
    BinStr = ""
    If Nbit < 1 Then MsgBox ("Warning(Dec2BinStr32Bit)!!! Decimal Number or number of Bit is wrong")
    If Nbit = 32 Then
        Nbit = 30
        If num < 0 Then
            BinStr = "1"
        Else
            BinStr = "0"
        End If
    End If
    For i = Nbit To 0 Step -1
        If num And (2 ^ i) Then
            BinStr = BinStr & "1"
        Else
            BinStr = BinStr & "0"
        End If
    Next
    Dec2BinStr32Bit = BinStr
'    Debug.Print BinStr
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function BinStr2HexStr(ByVal BinStr As String, ByVal HexBit As Long) As String
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "BinStr2HexStr"

    Dim i As Integer, j As Integer
    Dim BinStrLen As Long
    Dim HexMOD As Integer
    Dim HexStr As String
    Dim HexVal As String
    Dim HexLen As Long

    HexStr = ""
    
    BinStrLen = Len(BinStr)
    If (BinStrLen Mod (4)) > 0 Then
        HexLen = (BinStrLen \ 4) + 1
    Else
        HexLen = BinStrLen \ 4
    End If
    
    If HexBit > HexLen Then
        HexLen = HexBit
    End If

    HexMOD = HexLen * 4 - BinStrLen
    
    If HexMOD > 0 Then
        For i = 0 To HexMOD - 1
            BinStr = "0" & BinStr
        Next i
    End If

    For i = 0 To HexLen - 1
        If Mid(BinStr, i * 4 + 1, 4) = "0000" Then
            HexVal = "0"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0001" Then
            HexVal = "1"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0010" Then
            HexVal = "2"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0011" Then
            HexVal = "3"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0100" Then
            HexVal = "4"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0101" Then
            HexVal = "5"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0110" Then
            HexVal = "6"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0111" Then
            HexVal = "7"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1000" Then
            HexVal = "8"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1001" Then
            HexVal = "9"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1010" Then
            HexVal = "A"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1011" Then
            HexVal = "B"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1100" Then
            HexVal = "C"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1101" Then
            HexVal = "D"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1110" Then
            HexVal = "E"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1111" Then
            HexVal = "F"
        Else
            HexVal = "X"
        End If

        HexStr = HexStr & HexVal
    Next i

    BinStr2HexStr = HexStr
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Function Bin2Dec(sMyBin As String) As Long
    Dim x As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        Bin2Dec = Bin2Dec + Mid(sMyBin, iLen - x + 1, 1) * 2 ^ x
    Next
End Function

Function Bin2Dec_rev(sMyBin As String) As Variant
    Dim x As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        Bin2Dec_rev = Bin2Dec_rev + Mid(sMyBin, iLen - x + 1, 1) * 2 ^ (iLen - x)
    Next
End Function

Function Bin2Dec_rev_Double(sMyBin As String) As Double
    Dim x As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        Bin2Dec_rev_Double = Bin2Dec_rev_Double + Mid(sMyBin, iLen - x + 1, 1) * 2 ^ (iLen - x)
    Next
End Function


Public Function ExculdePath(Pat As Variant) As String
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "ExculdePath"

Dim patt_ary_temp() As String
    patt_ary_temp = Split(Pat, "\")
    ExculdePath = patt_ary_temp(UBound(patt_ary_temp))

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


''debug printing

Public Function DebugPrintFunc(Test_Pattern As String, Optional testname_enable As Boolean = False) As Long
'for debug printing generation
    Dim PinCnt As Long, pinary() As String
    Dim i As Long
    Dim PowerVolt As Double
    Dim Powerfoldlimit As Double
    Dim AlramCheck As String
    Dim PowerAlramTime As Double
    Dim All_power_list As PinList
    Dim CurrentChans As String
    Dim PatSetArray() As String
    Dim PrintPatSet As Variant
    Dim patt As Variant 'patt1
    Dim patt1 As Variant
    Dim patt_ary_debug() As String
    Dim pat_count_debug As Long
    Dim patt_ary_debug1() As String
    Dim pat_count_debug1 As Long
    Dim PinGroup() As String
    Dim EachPinGroup As Variant
    Dim Timelist As String
    Dim TimeGroup() As String
    Dim CurrTiming As Variant
    Dim TimeDomainlist As String
    Dim TimeDomaingroup() As String
    Dim CurrTimeDomain As Variant
    Dim TimeDomainIn As String
    Dim TempString As String
    Dim TempStringOffline As String
    Dim AlarmBehavior As tlAlarmBehavior
    Dim DebugPrint_version As Double
    Dim Vmain As Double
    Dim IRange As Double
    Dim Gate_State As Boolean
    Dim Gate_State_str As String
    Dim PinData As New PinListData
    Dim out_line As String
    Dim CurrSite As Variant
    Dim XI0_Vicm  As Double
    Dim XI0_Vid As Double
    Dim XI0_Vihd As Double
    Dim XI0_Vild As Double
    Dim SlotType As String
    
    On Error GoTo ErrHandler
    
    'version history
    'DebugPrint_version = 1.3   'copy from Fiji
    'DebugPrint_version = 1.4   'implement offline simulation for Rhea bring up
    'DebugPrint_version = 1.5   'Update for Multi-Port nWire setting
     DebugPrint_version = 1.6   'Add differential nWire frequency capture, DCVS tl* modes put in strings, support no pattern items
     DebugPrint_version = 1.7   'Add DC/AC cetegory setup, remove off-limt timing simulation, offline could get real timing.
     DebugPrint_version = 1.71   'Add PPMU debug print function.
     DebugPrint_version = 1.72   'Add DCVI debug print support.
     Shmoo_Pattern = Test_Pattern
     m1_InstanceName = LCase(TheExec.DataManager.InstanceName)
    
    'setups

    If DebugPrintFlag_Chk = True Then
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "================debug print start=================="
        'list all power pin's level
        TheExec.Datalog.WriteComment "  DebugPrint version = " & DebugPrint_version
        If testname_enable Then
            TheExec.Datalog.WriteComment "  TestInstanceName = " & G_TestName
            testname_enable = False
        Else
            TheExec.Datalog.WriteComment "  TestInstanceName = " & TheExec.DataManager.InstanceName
        End If
        
        TheExec.Datalog.WriteComment "***** List all Category info Start ******"
        ''''Get the current TestInstance Context
        Dim m_DCCategory As String
        Dim m_DCSelector As String
        Dim m_ACCategory As String
        Dim m_ACSelector As String
        Dim m_TimeSetSheet As String
        Dim m_EdgeSetSheet As String
        Dim m_LevelsSheet As String
        Dim m_tmpPMname As String
    
        ''''20151109
        ''''Use the local module private global variable to be flexible if it could be used anywhere in this Module. (Just in case)
        Call TheExec.DataManager.GetInstanceContext(m_DCCategory, m_DCSelector, _
                                                    m_ACCategory, m_ACSelector, _
                                                    m_TimeSetSheet, m_EdgeSetSheet, _
                                                    m_LevelsSheet, "")
    
        TempString = "DC Category ="
        TempString = TempString + " " + m_DCCategory
        TheExec.Datalog.WriteComment TempString
    
        TempString = "AC Category ="
        TempString = TempString + " " + m_ACCategory
        TheExec.Datalog.WriteComment TempString
    
        TempString = "Level ="
        TempString = TempString + " " + m_LevelsSheet
        TheExec.Datalog.WriteComment TempString

        TempString = "TimingSet ="
        TempString = TempString + " " + m_TimeSetSheet
        TheExec.Datalog.WriteComment TempString

        TheExec.Datalog.WriteComment "***** List all Category info end ******"
        TheExec.Datalog.WriteComment "***** List all power Start ******"

        TheExec.DataManager.DecomposePinList AllPowerPinlist, pinary(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        TempStringOffline = PinAry(i) & "_GLB"
'                        If LCase(TheExec.DataManager.InstanceName) Like "*_hv*" Then
'                            Vmain = TheExec.specs.Globals(TempStringOffline).ContextValue * TheExec.specs.Globals("Ratio_Plus").ContextValue
'                        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*_nv*" Then
'                            Vmain = TheExec.specs.Globals(TempStringOffline).ContextValue '* TheExec.Specs.Globals("Ratio_Plus").ContextValue
'                        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*_lv*" Then
'                            Vmain = TheExec.specs.Globals(TempStringOffline).ContextValue * TheExec.specs.Globals("Ratio_Minus").ContextValue
'                        End If
'                        'PowerVolt = Vmain
'                        PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Main.Value
'                    Else    'online
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerVolt = TheHdw.DCVS.Pins(pinary(i)).Voltage.Main.Value
                            Case "vhdvs": PowerVolt = TheHdw.DCVS.Pins(pinary(i)).Voltage.Main.Value
                            Case "dc-07": PowerVolt = TheHdw.DCVI.Pins(pinary(i)).Voltage
                        End Select
'                    End If
                        
                    TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & Format(PowerVolt, "0.000") & " v"
                End If
            Next i

            TheExec.Datalog.WriteComment "***** List all power end ******"
            
            TheExec.Datalog.WriteComment "***** List all Vmain power Start ******"

            TheExec.DataManager.DecomposePinList AllPowerPinlist, pinary(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerVolt = TheHdw.DCVS.Pins(pinary(i)).Voltage.Main.Value
                            Case "vhdvs": PowerVolt = TheHdw.DCVS.Pins(pinary(i)).Voltage.Main.Value
                            Case "dc-07": PowerVolt = TheHdw.DCVI.Pins(pinary(i)).Voltage
                        End Select
                    TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & Format(PowerVolt, "0.000") & " v"
                End If
            Next i

            TheExec.Datalog.WriteComment "***** List all Vmain power end ******"
        
            TheExec.Datalog.WriteComment "***** List all Valt power Start ******"

            TheExec.DataManager.DecomposePinList AllPowerPinlist, pinary(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerVolt = TheHdw.DCVS.Pins(pinary(i)).Voltage.Alt.Value
                            Case "vhdvs": PowerVolt = TheHdw.DCVS.Pins(pinary(i)).Voltage.Alt.Value
                            Case "dc-07": PowerVolt = TheHdw.DCVI.Pins(pinary(i)).Voltage
                        End Select
                    TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & Format(PowerVolt, "0.000") & " v"
                End If
            Next i

            TheExec.Datalog.WriteComment "***** List all Valt power end ******"
            
            TheExec.Datalog.WriteComment "***** List all power FoldLimit TimeOut Start ******"

            TempString = "FoldLimit TimeOut :"
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'PowerAlramTime = 0.001 * i
'                        PowerAlramTime = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.TimeOut
'                    Else    'online
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerAlramTime = TheHdw.DCVS.Pins(pinary(i)).CurrentLimit.Source.FoldLimit.TimeOut
                            Case "vhdvs": PowerAlramTime = TheHdw.DCVS.Pins(pinary(i)).CurrentLimit.Source.FoldLimit.TimeOut
                            Case "dc-07": PowerAlramTime = TheHdw.DCVI.Pins(pinary(i)).FoldCurrentLimit.TimeOut
                        End Select
'                    End If

                    If i <> (PinCnt - 1) Then
                        TempString = TempString + "  " & pinary(i) & " = " & Format(1000 * PowerAlramTime, "0.000") & " ms" + ","
                    Else
                        TempString = TempString + "  " & pinary(i) & " = " & Format(1000 * PowerAlramTime, "0.000") & " ms"
                    End If
                End If
            Next i

            TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power FoldLimit TimeOut End ******"
            TheExec.Datalog.WriteComment "***** List all power FoldLimit Current Start ******"

            TempString = "FoldLimit Current :"
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        TempStringOffline = PinAry(i) & "_Ifold_GLB"
'                        Irange = TheExec.specs.Globals(TempStringOffline).ContextValue
'                        'Powerfoldlimit = Irange
'                        Powerfoldlimit = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.Level.Value
'                    Else    'online
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": Powerfoldlimit = TheHdw.DCVS.Pins(pinary(i)).CurrentLimit.Source.FoldLimit.Level.Value
                            Case "vhdvs": Powerfoldlimit = TheHdw.DCVS.Pins(pinary(i)).CurrentLimit.Source.FoldLimit.Level.Value
                            Case "dc-07": Powerfoldlimit = TheHdw.DCVI.Pins(pinary(i)).Current
                        End Select
'                    End If

                    If i <> (PinCnt - 1) Then
                        TempString = TempString + "  " & pinary(i) & " = " & Format(Powerfoldlimit, "0.000000") & " A" + ","
                    Else
                        TempString = TempString + "  " & pinary(i) & " = " & Format(Powerfoldlimit, "0.000000") & " A"
                    End If
                End If
            Next i

            TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power FoldLimit Current End ******"
            TheExec.Datalog.WriteComment "***** List all power Alram Check Start ******"

            TempString = "Alram Check :"
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'AlarmBehavior = tlAlarmDefault
'                        AlarmBehavior = TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout)
'                    Else
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": AlarmBehavior = TheHdw.DCVS.Pins(pinary(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout)
                            Case "vhdvs": AlarmBehavior = TheHdw.DCVS.Pins(pinary(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout)
                            Case "dc-07": AlarmBehavior = TheHdw.DCVI.Pins(pinary(i)).FoldCurrentLimit.Behavior
                        End Select
'                    End If
                    
                    If AlarmBehavior = tlAlarmOff Then
                        AlramCheck = "tlAlarmOff"
                    ElseIf AlarmBehavior = tlAlarmContinue Then
                        AlramCheck = "tlAlarmContinue"
                    ElseIf AlarmBehavior = tlAlarmDefault Then
                        AlramCheck = "tlAlarmDefault"
                    ElseIf AlarmBehavior = tlAlarmForceBin Then
                        AlramCheck = "tlAlarmForceBin"
                    ElseIf AlarmBehavior = tlAlarmForceFail Then
                        AlramCheck = "tlAlarmForceFail"
                    End If
                    If i <> (PinCnt - 1) Then
                        TempString = TempString + "  " & pinary(i) & " = " & AlramCheck & ","
                    Else
                        TempString = TempString + "  " & pinary(i) & " = " & AlramCheck
                    End If
                End If
            Next i

            TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power Alram Check End ******"
            TheExec.Datalog.WriteComment "***** List all power Connection Check Start ******"

            TempString = "Power Relay Connection:"
            Dim PowerConnect_State As tlDCVSConnectWhat
            Dim PowerConnect_State_str As String
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'PowerConnect_State = tlDCVSConnectForce
'                        PowerConnect_State = TheHdw.DCVS.Pins(PinAry(i)).Connected
'                    Else
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerConnect_State = TheHdw.DCVS.Pins(pinary(i)).Connected
                            Case "vhdvs": PowerConnect_State = TheHdw.DCVS.Pins(pinary(i)).Connected
                            Case "dc-07": PowerConnect_State = TheHdw.DCVI.Pins(pinary(i)).Connected
                        End Select
'                    End If
                    
                    Select Case PowerConnect_State
                         Case tlDCVSConnectDefault: PowerConnect_State_str = "tlDCVSConnectDefault"
                         Case tlDCVSConnectNone: PowerConnect_State_str = "tlDCVSConnectNone"
                         Case tlDCVSConnectForce: PowerConnect_State_str = "tlDCVSConnectForce"
                         Case tlDCVSConnectSense: PowerConnect_State_str = "tlDCVSConnectSense"
                    End Select
                    If i <> (PinCnt - 1) Then
                        TempString = TempString + "  " & pinary(i) & " = " & PowerConnect_State_str + ","
                    Else
                        TempString = TempString + "  " & pinary(i) & " = " & PowerConnect_State_str
                    End If
                End If
            Next i

            TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power Connection Check End ******"
            TheExec.Datalog.WriteComment "***** List all power Gate Start ******"

            TempString = "Power Gate Status:"

            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'Gate_State = True
'                        Gate_State = TheHdw.DCVS.Pins(PinAry(i)).Gate
'                    Else
                       SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": Gate_State = TheHdw.DCVS.Pins(pinary(i)).Gate
                            Case "vhdvs": Gate_State = TheHdw.DCVS.Pins(pinary(i)).Gate
                            Case "dc-07": Gate_State = TheHdw.DCVI.Pins(pinary(i)).Gate
                        End Select
'                    End If
                    
                    Select Case Gate_State
                         Case True: Gate_State_str = "on"
                         Case False: Gate_State_str = "off"
                    End Select
                    If i <> (PinCnt - 1) Then
                        TempString = TempString + "  " & pinary(i) & " = " & Gate_State_str + ","
                    Else
                        TempString = TempString + "  " & pinary(i) & " = " & Gate_State_str
                    End If
                End If
            Next i

            TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power Gate Check End ******"
            TheExec.Datalog.WriteComment "***** List Pattern Start ******"

            'Print test pattern
            If Test_Pattern <> "" Then
                PatSetArray = Split(Test_Pattern, ",")

                For Each PrintPatSet In PatSetArray
                    If LCase(PrintPatSet) Like "*.pat*" Then
                        TheExec.Datalog.WriteComment "  Pattern : " & PrintPatSet
                    Else
                        GetPatListFromPatternSet CStr(PrintPatSet), patt_ary_debug, pat_count_debug
                        For Each patt In patt_ary_debug
                            If patt <> "" Then TheExec.Datalog.WriteComment "  Pattern : " & patt
                        Next patt
                    End If
                Next PrintPatSet
            Else
                'do nothing, no printing
            End If

            TheExec.Datalog.WriteComment "***** List Pattern end ******"
            TheExec.Datalog.WriteComment "***** List Level Start ******"

            PinGroup = Split(PinGrouplist, ",")
'====================================================================Mask differential pins start====================================================================
            Dim Pins() As String
            Dim Diff_pins() As String
            Dim pincont As Long
            Dim pincont_diff As Long
            Dim Diffenential_Pins As Variant
            Dim I_diff As Integer
            Dim J_original As Integer
            
            Diffenential_Pins = "All_DiffPairs"
            For Each EachPinGroup In PinGroup   'EachPinGroup
                TheExec.DataManager.DecomposePinList EachPinGroup, Pins, pincont
                TheExec.DataManager.DecomposePinList Diffenential_Pins, Diff_pins, pincont_diff
                If DicDiffPairs.Exists(Diff_pins(0)) = False Then
                    For I_diff = 0 To pincont_diff - 1
                        DicDiffPairs.Add Diff_pins(I_diff), Diff_pins(I_diff)
                    Next I_diff
                End If
                For J_original = 0 To pincont - 1
                   If DicDiffPairs.Exists(Pins(J_original)) = False Then
                        If TheExec.DataManager.ChannelType(Pins(J_original)) <> "N/C" Then
                           TheExec.Datalog.WriteComment "  Pins : " & CStr(EachPinGroup) _
                           & " , Vih = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.Value(chVih), "0.000") & " v" _
                           & " , Vil = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.Value(chVil), "0.000") & " v" _
                           & " , Voh = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.Value(chVoh), "0.000") & " v" _
                           & " , Vol = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.Value(chVol), "0.000") & " v" _
                           & " , Iol = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.Value(chIoh), "0.000") & " v" _
                           & " , Ioh = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.Value(chIol), "0.000") & " v" _
                           & " , Vt  = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.Value(chVt), "0.000") & " v" _
                           & " , Vch = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.Value(chVch), "0.000") & " v" _
                           & " , Vcl = " & Format(TheHdw.Digital.Pins(CStr(Pins(J_original))).Levels.Value(chVcl), "0.000") & " v" _
                           & " , PPMU_VclampHi = " & Format(TheHdw.PPMU.Pins(CStr(Pins(J_original))).ClampVHi, "0.000") & " v" _
                           & " , PPMU_VclampLow = " & Format(TheHdw.PPMU.Pins(CStr(Pins(J_original))).ClampVLo, "0.000") & " v"
                           'Diff_pins_dictionary.RemoveAll
                           Exit For
                        End If
                   End If
                Next J_original
'====================================================================Mask differential pins end====================================================================

                TheExec.Datalog.WriteComment "  Pins : " & CStr(EachPinGroup) _
                & " , Vih = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVih), "0.000") & " v" _
                & " , Vil = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVil), "0.000") & " v" _
                & " , Voh = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVoh), "0.000") & " v" _
                & " , Vol = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVol), "0.000") & " v" _
                & " , Iol = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chIoh), "0.000") & " v" _
                & " , Ioh = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chIol), "0.000") & " v" _
                & " , Vt  = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVt), "0.000") & " v" _
                & " , Vch = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVch), "0.000") & " v" _
                & " , Vcl = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVcl), "0.000") & " v" _
                & " , PPMU_VclampHi = " & Format(TheHdw.PPMU.Pins(CStr(EachPinGroup)).ClampVHi, "0.000") & " v" _
                & " , PPMU_VclampLow = " & Format(TheHdw.PPMU.Pins(CStr(EachPinGroup)).ClampVLo, "0.000") & " v"
            Next EachPinGroup

            TheExec.Datalog.WriteComment "***** List Level end ******"
            TheExec.Datalog.WriteComment "***** List Timing Start ******"

            If Test_Pattern <> "" Then
                TimeDomainlist = TheHdw.Digital.Timing.TimeDomainlist
                TimeDomaingroup = Split(TimeDomainlist, ",")
                For Each CurrTimeDomain In TimeDomaingroup
                    If CStr(CurrTimeDomain) = "All" Then
                        TimeDomainIn = ""
                    Else
                        TimeDomainIn = CStr(CurrTimeDomain)
                    End If
                    
                    Timelist = TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
                    'TimeGroup
                    TimeGroup = Split(Timelist, ",")
                    For Each CurrTiming In TimeGroup
                        If CurrTiming <> "" Then
                            If TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.Period(CStr(CurrTiming)) > 0 Then
                                TheExec.Datalog.WriteComment "  Time Doamin : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & Format((1 / TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.Period(CStr(CurrTiming))) / 1000000, "0.000") & " Mhz"
                            Else
                                TheExec.Datalog.WriteComment "  Time Doamin : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & Format(0, "0.000") & " Mhz"
                            End If
                        End If
                    Next CurrTiming
                Next CurrTimeDomain
            Else
                TheExec.Datalog.WriteComment "  Time Doamin : " & "N/A" & ", TimeSet : " & "N/A" & " = " & Format(0 / 1000000, "0.000") & " Mhz"
            End If

            '' add for XI0 free running clk
'               TheExec.Datalog.WriteComment "  FreeRunFreq : " & TheHdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz , clock_Vih: " & TheHdw.DIB.SupportBoardClock.Vih & " v , clock_Vil: " & TheHdw.DIB.SupportBoardClock.Vil & " v"
            Dim XI0_Freq_pl As New PinListData, RTCLK_Freq_pl As New PinListData, Pin_XI0 As New PinList, Pin_RTCLK As New PinList
            Dim Site As Variant
          
            If XI0_GP <> "" Then 'differential(false) or single end(true)
                Pin_XI0.Value = XI0_GP
                TheHdw.Digital.Pins(Pin_XI0).Levels.Value(chVoh) = TheHdw.Digital.Pins(Pin_XI0).Levels.Value(chVih) / 4
            ElseIf XI0_Diff_GP <> "" Then
                'Vod=0, do nothing
                Pin_XI0.Value = XI0_Diff_GP
                TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVod) = TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVid) / 4
            End If

            If XI0_Diff_GP <> "" Or XI0_GP <> "" Then
                Freq_MeasFreqSetup Pin_XI0, 0.001
                Freq_MeasFreqStart Pin_XI0, 0.001, XI0_Freq_pl
            End If

            If TheExec.TesterMode = testModeOffline Then
                For Each Site In TheExec.Sites
                    XI0_Freq_pl.Pins(0).Value = 24000000
                Next Site
            End If

            For Each Site In TheExec.Sites
                    If XI0_GP <> "" Then 'differential(false) or single end(true)
                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0) : " & Format(XI0_Freq_pl.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(Pin_XI0).Levels.Value(chVih), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(Pin_XI0).Levels.Value(chVil), "0.000") & " v"
                        'CHWu modify 10/14 to add Xio_PA_1 and remove RTCLK
'                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0_1) : " & Format(XI0_Freq_pl_1.pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.pins(Pin_XI0_1).Levels.Value(chVih), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.pins(Pin_XI0_1).Levels.Value(chVil), "0.000") & " v"
                    ElseIf XI0_Diff_GP <> "" Then
                      'CHWu modify 11/17 modify for Xio_PA printout
                       XI0_Vicm = TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVicm)
                       XI0_Vid = TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVid)
                       XI0_Vihd = XI0_Vicm + XI0_Vid / 2
                       XI0_Vild = XI0_Vicm - XI0_Vid / 2
                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0) : " & Format(XI0_Freq_pl.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(XI0_Vihd, "0.000") & " v , clock_Vil: " & Format(XI0_Vild, "0.000") & " v"
'                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0_1) : " & Format(XI0_Freq_pl_1.pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(XI0_Vihd, "0.000") & " v , clock_Vil: " & Format(XI0_Vild, "0.000") & " v"
'                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0) : " & Format(XI0_Freq_pl.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVid), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVod), "0.000") & " v"
'                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0_1) : " & Format(XI0_Freq_pl_1.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(Pin_XI0_1).DifferentialLevels.Value(chVid), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(Pin_XI0_1).DifferentialLevels.Value(chVod), "0.000") & " v"
                    End If
            Next Site

            Meas_FRC ""    ' Multi nWire 20170718

            TheExec.Datalog.WriteComment "***** List Timing end ******"
            TheExec.Datalog.WriteComment "***** List Disable Compare check Start ******"
            
            'EachPinGroup
            PinGroup = Split(PinGrouplist, ",")
            For Each EachPinGroup In PinGroup
                TheExec.DataManager.DecomposePinList EachPinGroup, Pins, pincont
                For J_original = 0 To pincont - 1
                    If TheExec.DataManager.ChannelType(Pins(J_original)) <> "N/C" Then
                        TheExec.Datalog.WriteComment "  Pins : " & CStr(Pins(J_original)) _
                    & " , Disable Compare= " & TheHdw.Digital.Pins(Pins(J_original)).DisableCompare
                    Exit For
                    End If
                Next J_original
            Next EachPinGroup

            TheExec.Datalog.WriteComment "***** List List Disable Compare check End ******"
            TheExec.Datalog.WriteComment "***** List all utility bit status Start ******"
            TheExec.DataManager.DecomposePinList All_Utility_list, pinary(), PinCnt

            'Utility bits
            out_line = "Utility_list : "
            For Each CurrSite In TheExec.Sites.Active
                For i = 0 To PinCnt - 1
                    If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
                        PinData = TheHdw.Utility.Pins(pinary(i)).States(tlUBStateProgrammed)    'TheHdw.Utility.pins((pinary(i)) '.States(tlUBStateCompared)
                        If i = 0 Then
                              out_line = out_line + pinary(i) & " = " & PinData.Pins(0).Value(CurrSite) '''& ","
                        Else
                              out_line = out_line & "," & pinary(i) & " = " & PinData.Pins(0).Value(CurrSite)
                        End If
                    End If
                Next i
                TheExec.Datalog.WriteComment out_line
                out_line = "Utility_list : "
            Next CurrSite

            TheExec.Datalog.WriteComment "***** List all utility bit status end ******"
            TheExec.Datalog.WriteComment "================debug print end  =================="
            TheExec.Datalog.WriteComment ""
        End If

    Exit Function
    
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

'Public Function Record_nWireScope(ByRef nWireScope As SiteBoolean, ByRef nWireScopeAllSite As Boolean, Optional DebugFlag As Boolean = False)
'    On Error GoTo ErrHandler
'    Dim TempStr1, TempStr2 As String
'    Dim Site As Variant
'    'Debug.Print TheExec.Sites.Active.Count
'    TempStr1 = Clock_Port
'    TempStr2 = Clock_Port1
'    'TheHdw.Protocol.Ports(TempStr).Halt
'    'check nWire status
'    For Each Site In TheExec.Sites.Active
'        nWireScope(Site) = False
'        If TheHdw.Protocol.ports(TempStr1).Enabled = False And TheHdw.Protocol.ports(TempStr2).Enabled = False Then
'            nWireScope(Site) = False
'            nWireScopeAllSite = False
'        ElseIf TheHdw.Protocol.ports(TempStr1).Enabled = True And TheHdw.Protocol.ports(TempStr2).Enabled = True Then
'            nWireScope(Site) = True
'            nWireScopeAllSite = True
'        Else
'            TheExec.Datalog.WriteComment "print: nWire dual-port not align, please check!!!"
'            'Stop
'            nWireScope(Site) = False
'        End If
'
'        If DebugFlag = True Then
'            TheExec.Datalog.WriteComment "print: Site(" & Site & ") nWire initial status, port(" & TempStr1 & "," & TempStr2 & ") " & nWireScope(Site)
'        End If
'    Next Site
'
'    Exit Function
'
'ErrHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function FreeRunClk_ScopeOut(PAPort As PinList, Optional DebugFlag As Boolean = False)
'
'    On Error GoTo ErrHandler
'    Dim TempStr As String
'
'    TheHdw.Digital.Pins(PAPort).Disconnect
'
'    If DebugFlag = True Then
'        TheExec.Datalog.WriteComment "print: nWire scope out, port (" & PAPort.Value & ")"
'    End If
'
'    Exit Function
'
'ErrHandler:
'    If AbortTest Then Exit Function Else Resume Next
'
'
'End Function
Public Function FreeRunClk_ScopeIn(PAPort As PinList, Optional DebugFlag As Boolean = False) ''update for multi nWire 20170718

    On Error GoTo ErrHandler
    Dim TempStr As String
                TheHdw.Digital.Pins(PAPort).Connect    '20190416
    Call Enable_FRC(PAPort.Value, True)
    TheHdw.Wait 0.001
    
'    TheHdw.Digital.Pins(PAPort).Connect
''
''    TheHdw.Protocol.ports(PAPort).Enabled = True
''    TheHdw.Protocol.ports(PAPort).NWire.ResetPLL
''    TheHdw.Wait 0.001
''    ' Start the nWire engine.
''    Call TheHdw.Protocol.ports(PAPort).NWire.Frames("RunFreeClock").Execute
''
    If DebugFlag = True Then
        TheExec.Datalog.WriteComment "print: nWire scope in, port (" & PAPort.Value & ")"
    End If
    
    Exit Function
    
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function

Public Function PowerUp_Interpose(PAPort As PinList, Optional DebugFlag As Boolean = False)
    On Error GoTo ErrHandler
    
'    FreeRunClk_ScopeOut PAPort, DebugFlag
    'TheHdw.Utility.Pins(Relay).State = tlUtilBitOn
    
'    If DebugFlag = True Then    'debugprint
'         TheExec.Datalog.WriteComment "print: RTCLK relay on, relay " & Relay.Value
'    End If

    FreeRunClk_ScopeIn PAPort, DebugFlag
    Exit Function

ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function PowerDown_Interpose(nWireDisconnectPin As String, Optional DebugFlag As Boolean = False)
    On Error GoTo ErrHandler

    FreeRunClk_Disconnect nWireDisconnectPin, DebugFlag
    
    TheExec.Datalog.WriteComment "print: nWire engine , Halt " & vbCrLf

    Exit Function
    
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

'Public Function AddNum()
'On Error GoTo ErrHandler
'Dim funcName As String:: funcName = "AddNum"
'
'Dim x As Integer
'Dim y As Integer
'
'Dim InitNum As Long
'Dim idx As Long
'InitNum = 130000
'
'
'
'idx = 0
'For x = 5 To 10000
'
'
'    If LCase(Cells(x, 7)) Like "test" Or LCase(Cells(x, 7)) Like "characterize" Then
'
'        Cells(x, 10) = InitNum + 100 * idx
'
'    idx = idx + 1
'
'    End If
'
'    If Cells(x, 7) = "" Then x = 10000
'
'Next x
'
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function
Function IEDA_Initialize(ByRef InputStr As String)

On Error GoTo ErrHandler
    Dim funcName As String:: funcName = "IEDA_Initialize"
    
    InputStr = ""

Exit Function

ErrHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

'20190416 top
Function IEDA_GetString(ByRef InputStr As String, RegistryName As String)
'(ByRef InputStr As String, FuseCategory As String, CategoryIndex As Integer)
On Error GoTo ErrHandler
    Dim funcName As String:: funcName = "IEDA_GetString"
    Dim Site As Variant
    Dim TmpString As String

        For Each Site In TheExec.Sites.Existing
                        Select Case RegistryName
                        'ECID IEDA
'                           Case "eFuseLotNumber"
'                                TmpString = ECIDFuse.Category(ECIDIndex("Lot_ID")).Read.ValStr(Site)
'                            Case "eFuseWaferID"
'                                TmpString = ECIDFuse.Category(ECIDIndex("Wafer_ID")).Read.ValStr(Site)
'                            Case "eFuseDieX"
'                                TmpString = ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.ValStr(Site)
'                            Case "eFuseDieY"
'                                TmpString = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.ValStr(Site)
'                            Case "Hram_ECID_53bit"
'                                TmpString = ECIDFuse.Category(gI_Index_DEID).Read.BitStrL(Site)
                            
''                            If CategoryIndex = gI_Index_53bits Then
''                                TmpString = ECIDFuse.Category(CategoryIndex).Read.BitStrL(Site)
''                            Else
''                                TmpString = ECIDFuse.Category(CategoryIndex).Read.ValStr(Site)
''                            End If
                          'UID IEDA
                        Case "Prov_Code"
'                                Call IEDA_UID_Decode
'                                If UIDFuse.Category(UIDIndex("UID_Code")).Read.ValStr(Site) = "" Then
'                                    TmpString = ""  'site not enable
'                                ElseIf CDbl(UIDFuse.Category(UIDIndex("UID_Code")).Read.ValStr(Site)) = 0 Then
'                                    TmpString = "0"
'                                ElseIf CDbl(UIDFuse.Category(UIDIndex("UID_Code")).LoLMT) < CDbl(UIDFuse.Category(UIDIndex("UID_Code")).Read.ValStr(Site)) And CDbl(UIDFuse.Category(UIDIndex("UID_Code")).Read.ValStr(Site)) < CDbl(UIDFuse.Category(UIDIndex("UID_Code")).HiLMT) Then
'                                    TmpString = "1"
'                                End If
                        'CFG IEDA
                        Case "SVM_CFuse"
'                            TmpString = CFGFuse.Category(gI_CFG_firstbits_index).Read.BitStrM(Site)
''                            If CategoryIndex = gI_CFG_firstbits_index Then
''                                TmpString = CFGFuse.Category(CategoryIndex).Read.BitStrM(Site)
''                            Else
''                                TmpString = CFGFuse.Category(CategoryIndex).Read.ValStr(Site)
''                            End If
                        Case "TMPS1_Untrim"
                            TmpString = gS_TMPS1_Untrim(Site)
                        Case "TMPS2_Untrim"
                            TmpString = gS_TMPS2_Untrim(Site)
                        Case "TMPS3_Untrim"
                            TmpString = gS_TMPS3_Untrim(Site)
                        Case "TMPS4_Untrim"
                            TmpString = gS_TMPS4_Untrim(Site)
                        Case "TMPS5_Untrim"
                            TmpString = gS_TMPS5_Untrim(Site)
                        Case "TMPS6_Untrim"
                            TmpString = gS_TMPS6_Untrim(Site)
                        Case "TMPS7_Untrim"
                            TmpString = gS_TMPS7_Untrim(Site)
                        Case "TMPS8_Untrim"
                            TmpString = gS_TMPS8_Untrim(Site)
                        
                        Case "TMPS1_Trim"
                            TmpString = gS_TMPS1_Trim(Site)
                        Case "TMPS2_Trim"
                            TmpString = gS_TMPS2_Trim(Site)
                        Case "TMPS3_Trim"
                            TmpString = gS_TMPS3_Trim(Site)
                        Case "TMPS4_Trim"
                            TmpString = gS_TMPS4_Trim(Site)
                        Case "TMPS5_Trim"
                            TmpString = gS_TMPS5_Trim(Site)
                        Case "TMPS6_Trim"
                            TmpString = gS_TMPS6_Trim(Site)
                        Case "TMPS7_Trim"
                            TmpString = gS_TMPS7_Trim(Site)
                        Case "TMPS8_Trim"
                            TmpString = gS_TMPS8_Trim(Site)
''                        Case "UDR"
''                            TmpString = UDRFuse.Category(CategoryIndex).Read.ValStr(Site)
''                        Case "SEN"
''                            TmpString = SENFuse.Category(CategoryIndex).Read.ValStr(Site)

                        Case Else
                            TheExec.Datalog.WriteComment "print: warnining, no suitable registry choosed in VBT 'IEDA_GetString'."
                        End Select
            If (Site = TheExec.Sites.Existing.Count - 1) Then
                InputStr = InputStr + TmpString
            Else
                 InputStr = InputStr + TmpString + ","
            End If
            
        Next Site

Exit Function

ErrHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Function IEDA_AutoCheck_Print(ByRef InputStr As String, RegistryName As String, DebugPrint As Boolean)

On Error GoTo ErrHandler
    Dim funcName As String:: funcName = "IEDA_AutoCheck_Print"
    Dim TmpString As String

    'InputStr = auto_checkIEDAString(InputStr)
    If DebugPrint Then TheExec.Datalog.WriteComment "print: Set IEDA registry ( " & RegistryName & " ) = " & InputStr
    
Exit Function

ErrHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function
'20190416 end

Function IEDA_SaveRegistry(ByVal InputStr As String, RegistryName As String)

On Error GoTo ErrHandler
    Dim funcName As String:: funcName = "IEDA_SaveRegistry"

    Call RegKeySave(RegistryName, InputStr)
    
Exit Function

ErrHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

'Public Function SetFRCPath(PinName As String)
'
'    ' Setup IKS mask for PIK & PGIK
'    Call TheHdw.Raw.TSIO.Wr("A_GL_IK_MSK", 9)
'
'    ' Set the steering to select the pin
'    Dim PinKey As Long
'    Call m_stdsvcclient.IkSvc.GetPinKey(PinName, PinKey)
'    Call m_stdsvcclient.IkSvc.SelectPinKey(PinKey)
'
'    ' Bypass the DDR flop
'    Call TheHdw.Raw.TSIO.Wr("A_P_TIM_PA_DRV_SEL", 6)
'
'End Function
Public Function Bintable_initial()
'//=====================================================================================
    
On Error GoTo ErrHandler

    Dim BintableSheet As Worksheet
    Dim BinNameColumnMax As Long
    Dim BinColumnNum As Long
    Dim BinContext As String
    Dim BinColNumAccu As Long: BinColNumAccu = 0
    Dim m_bintable_exist_Flag As Boolean
    
    m_bintable_exist_Flag = False
    
    '20170529 add sheet loop to include all bin table sheets
    For Each BintableSheet In ThisWorkbook.Sheets
        If LCase(BintableSheet.Name) Like "*bin*table*" Then
            m_bintable_exist_Flag = True
            #If IGXL8p30 Then
            #Else
                BintableSheet.Activate
            #End If
            BinNameColumnMax = BintableSheet.Cells(Rows.Count, 2).End(xlUp).Row
            BinNameColumnMax = Worksheets(BintableSheet.Name).UsedRange.Rows.Count
            ReDim Preserve tyBinTable.astrBinName(BinNameColumnMax - 4 + BinColNumAccu)
            ReDim Preserve tyBinTable.astrBinRename(BinNameColumnMax - 4 + BinColNumAccu)
            ReDim Preserve tyBinTable.astrBinSortNum(BinNameColumnMax - 4 + BinColNumAccu)
            
            For BinColumnNum = 4 To BinNameColumnMax
                BinContext = BintableSheet.Cells(BinColumnNum, 2).Value
                If BinContext <> "" Then
                    tyBinTable.astrBinName(BinColumnNum - 4 + BinColNumAccu) = BinContext
                    tyBinTable.astrBinRename(BinColumnNum - 4 + BinColNumAccu) = BintableSheet.Cells(BinColumnNum, 1).Value
                    tyBinTable.astrBinSortNum(BinColumnNum - 4 + BinColNumAccu) = BintableSheet.Cells(BinColumnNum, 5).Value
                Else
                    Exit For
                End If
            Next BinColumnNum
        
            BinColNumAccu = BinColumnNum - 4 + BinColNumAccu
        End If
    Next BintableSheet
    
    Dim lCount As Long
    
    If (m_bintable_exist_Flag) Then
        For lCount = 0 To UBound(tyBinTable.astrBinName)
            If tyBinTable.astrBinSortNum(lCount) <> "" Then
                If tyBinTable.astrBinRename(lCount) <> "" Then
                    Call TheExec.Datalog.SBRFill(tyBinTable.astrBinSortNum(lCount), tyBinTable.astrBinRename(lCount))
                Else
                    Call TheExec.Datalog.SBRFill(tyBinTable.astrBinSortNum(lCount), Mid(tyBinTable.astrBinName(lCount), InStr(tyBinTable.astrBinName(lCount), "_") + 1))
                End If
            End If
        Next lCount
    End If

    Exit Function

ErrHandler:
    '//============================================================================================
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DebugPrintFunc_PPMU(PPMU_Pins As String) As Long
'for debug printing generation
    Dim PinCnt As Long, pinary() As String
    Dim i As Long
    Dim PowerVolt As Double
    Dim PowerCurrent As Double
    Dim Powerfoldlimit As Double
    Dim AlramCheck As String
    Dim PowerAlramTime As Double
    Dim All_power_list As PinList
    Dim PinGroup() As String
    Dim EachPinGroup As Variant
    Dim DebugPrint_version As Double
    Dim Pins() As String, Pin_Cnt As Long
    Dim PPMU_used_Pin As Variant
    Dim PPMU_ForceV As String
    Dim PPMU_ForceI As String
    Dim DCVI_Mode As String
    Dim DCVI_sense_relay As Boolean
    Dim DCVI_force_relay As Boolean
    
    On Error GoTo ErrHandler
    
    'version history
    DebugPrint_version = 1.71   'implement PPMU debug print function
    
    'setups

    If DebugPrintFlag_Chk = True Then
        
        If PPMU_Pins <> "" Then
        
            TheExec.DataManager.DecomposePinList PPMU_Pins, Pins(), Pin_Cnt
        Else
            Pin_Cnt = 0
        End If
        
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "================debug print PPMU start=================="
        'list all power pin's level
        TheExec.Datalog.WriteComment "  DebugPrint version = " & DebugPrint_version
        TheExec.Datalog.WriteComment "  TestInstanceName = " & TheExec.DataManager.InstanceName
        TheExec.Datalog.WriteComment "***** List all power Start ******"

        TheExec.DataManager.DecomposePinList AllPowerPinlist, pinary(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then

                        PowerVolt = TheHdw.DCVS.Pins(pinary(i)).Voltage.Main.Value
                        
                    TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & Format(PowerVolt, "0.000") & " v"
                End If
            Next i
        TheExec.Datalog.WriteComment "***** List all power end ******"



        TheExec.Datalog.WriteComment "***** List all DCVI Start ******"

        TheExec.DataManager.DecomposePinList AllDCVIPinlist, pinary(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then

                    PowerVolt = TheHdw.DCVI.Pins(pinary(i)).Voltage
                    PowerCurrent = TheHdw.DCVI.Pins(pinary(i)).Current
                    
                    If TheHdw.DCVI.Pins(pinary(i)).Mode = tlDCVIModeVoltage Then
                        DCVI_Mode = "ForceV"
                    ElseIf TheHdw.DCVI.Pins(pinary(i)).Mode = tlDCVIModeCurrent Then
                        DCVI_Mode = "ForceI"
                    Else
                        DCVI_Mode = "HighImpedance"
                    End If


                    If TheHdw.DCVI.Pins(pinary(i)).Connected = 0 Then
                        DCVI_force_relay = False
                        DCVI_sense_relay = False
                    ElseIf TheHdw.DCVI.Pins(pinary(i)).Connected = 1 Then
                        DCVI_force_relay = True
                        DCVI_sense_relay = False
                    ElseIf TheHdw.DCVI.Pins(pinary(i)).Connected = 2 Then
                        DCVI_force_relay = False
                        DCVI_sense_relay = True
                    ElseIf TheHdw.DCVI.Pins(pinary(i)).Connected = 3 Then
                        DCVI_force_relay = True
                        DCVI_sense_relay = True
                    End If

                    TheExec.Datalog.WriteComment "  DCVI_Pins : " & pinary(i) _
                    & " , Voltage = " & Format(PowerVolt, "0.000000") & " v" _
                    & " , Current = " & Format(PowerCurrent, "0.000000") & " A" _
                    & " , Mode = " & DCVI_Mode & " " _
                    & " , Gate = " & TheHdw.DCVI.Pins(pinary(i)).Gate _
                    & " , DCVI_sense_relay = " & DCVI_sense_relay _
                    & " , DCVI_force_relay = " & DCVI_force_relay
                    
                End If
            Next i
            
            
            
        TheExec.Datalog.WriteComment "***** List all DCVI end ******"
            
        TheExec.Datalog.WriteComment "***** List PPMU condition Start ******"
            
            


            
            If Pin_Cnt > 0 Then
                For Each PPMU_used_Pin In Pins
                    If TheExec.DataManager.ChannelType(PPMU_used_Pin) <> "N/C" Then
                        PPMU_ForceV = CStr(Format(TheHdw.PPMU.Pins(PPMU_used_Pin).Voltage.Value, "0.000000"))
                        PPMU_ForceI = CStr(Format(TheHdw.PPMU.Pins(PPMU_used_Pin).Current.Value, "0.000000"))
                        
                        If TheHdw.PPMU.Pins(CStr(PPMU_used_Pin)).Mode = tlPPMUForceVMeasureI Then
                            PPMU_ForceI = "None"
                        Else
                            PPMU_ForceV = "None"
                        End If
        
                        TheExec.Datalog.WriteComment "  Pins : " & CStr(PPMU_used_Pin) _
                        & " , PPMU_VclampHi = " & Format(TheHdw.PPMU.Pins(CStr(PPMU_used_Pin)).ClampVHi, "0.000") & " v" _
                        & " , PPMU_VclampLow = " & Format(TheHdw.PPMU.Pins(CStr(PPMU_used_Pin)).ClampVLo, "0.000") & " v" _
                        & " , PPMU_forceV = " & PPMU_ForceV & " v" _
                        & " , PPMU_ForceI = " & PPMU_ForceI & " A"
                  End If
                Next PPMU_used_Pin
            End If
        
        TheExec.Datalog.WriteComment "***** List PPMU condition end ******"
            

            TheExec.Datalog.WriteComment "================debug print PPMU end  =================="
            TheExec.Datalog.WriteComment ""
        End If
    Exit Function
    
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function GetPatFromPatternSet(TestPat As String, _
                              rtnPatNames() As String, _
                              rtnPatCnt As Long) As Boolean

    Dim PatCnt As Long                          '<- Number of patterns in set
    Dim RawNameData() As String                 '<- Raw pattern name data
    Dim rtnPatNames1() As String
    Dim rtnPatNames2() As String
    Dim i As Long, j As Long
    '___ Init _____________________________________________________________________________
    On Error GoTo ErrHandler
    
    '___ Check the name ___________________________________________________________________
    '    Individual pattern name or non-pattern string returns an error - thus false
    '--------------------------------------------------------------------------------------
    rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(TestPat, PatCnt)
    If (UBound(rtnPatNames) > 0) Then
        If LCase(rtnPatNames(0)) Like "*.pat*" Then
            GetPatFromPatternSet = True
            rtnPatCnt = UBound(rtnPatNames) + 1
        Else
            rtnPatCnt = 0
            For i = 0 To UBound(rtnPatNames)
                rtnPatNames2 = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames(i), PatCnt)
                rtnPatCnt = rtnPatCnt + UBound(rtnPatNames2) + 1
            Next i
            rtnPatNames1 = TheExec.DataManager.Raw.GetPatternsInSet(TestPat, PatCnt)
            ReDim rtnPatNames(rtnPatCnt)
            rtnPatCnt = 0
            For i = 0 To UBound(rtnPatNames1)
                rtnPatNames2 = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames1(i), PatCnt)
                For j = 0 To UBound(rtnPatNames2)
                    If LCase(rtnPatNames2(j)) Like "*.pat*" Then
                        rtnPatNames(rtnPatCnt) = rtnPatNames2(j)
                    Else
                        TheExec.ErrorLogMessage TestPat & " in more than 2 level of pattern set"
                    End If
                    rtnPatCnt = rtnPatCnt + 1
                Next j
            Next i
            GetPatFromPatternSet = True
        End If
    Else
        If LCase(rtnPatNames(0)) Like "*.pat*" Then
            GetPatFromPatternSet = True
            rtnPatCnt = 1
        Else
            rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames(0), PatCnt)
            rtnPatCnt = UBound(rtnPatNames) + 1
            For j = 0 To UBound(rtnPatNames)
                If LCase(rtnPatNames(j)) Like "*.pat*" Then
                Else
                    TheExec.ErrorLogMessage TestPat & " in more than 2 level of pattern set"
                End If
            Next j
        End If
    End If
    
    Exit Function
    
Exit Function
ErrHandler:
    GetPatFromPatternSet = False
    rtnPatCnt = -1

                If AbortTest Then Exit Function Else Resume Next
End Function
Public Function FreeRunClk_Disconnect(nWireDisconnectPin As String, Optional DebugFlag As Boolean = False)

On Error GoTo ErrHandler
    Dim funcName As String:: funcName = "FreeRunClk_Disconnect"

    TheHdw.Digital.Pins(nWireDisconnectPin).Disconnect
    
    If DebugFlag = True Then TheExec.Datalog.WriteComment "print: nWire disconnect, pin" & nWireDisconnectPin

    Exit Function
    
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'Public Function Find_nWire_Pin() As Long  ''Support multiple nWire port 20170718
'On Error GoTo ErrorHandler
'' Get all nWire port and put in global variable nWire_Ports_GLB
'    Dim i As Long
'    Dim ws As Worksheet
'    Dim wb As Workbook
'    Dim row_cnt As Long
'    Dim nWire_cnt As Long
'    Dim nWire_Pin_ary(10) As String
'    Dim curr_pin As String, last_pin As String
'
''    If nWire_Ports_GLB <> "" Then Exit Function
'    nWire_Ports_GLB = ""
'
'    Set wb = Application.ActiveWorkbook
'
'    Dim m_sheetName As String
'    m_sheetName = "Levels_nWire"
'    If (CheckIfSheetExists(m_sheetName) = False) Then Exit Function
'
'    Set ws = wb.Sheets(m_sheetName)
'    ws.Activate
'
''    row_cnt = ws.Cells(Rows.Count, 2).End(xlUp).Row - 1
'    nWire_cnt = 0
'    For i = 4 To Rows.Count 'skip header line
'        curr_pin = ws.Cells(i, 2)
'        If curr_pin = "" Then
'            i = Rows.Count + 1 'stop at empty row/cell
'        ElseIf curr_pin Like "*_PA" And curr_pin <> last_pin Then
'            nWire_Pin_ary(nWire_cnt) = curr_pin
'            last_pin = curr_pin
'            nWire_cnt = nWire_cnt + 1
'        End If
'    Next i
'    For i = 1 To nWire_cnt
'        If nWire_Ports_GLB <> "" Then
'            nWire_Ports_GLB = nWire_Ports_GLB & "," & nWire_Pin_ary(i - 1)
'        Else
'            nWire_Ports_GLB = nWire_Pin_ary(i - 1)
'        End If
'    Next i
'
'    Exit Function
'ErrorHandler:
'    TheExec.AddOutput "<Error> Find_nWire_Pin:: please check it out."
'    TheExec.Datalog.WriteComment "<Error> Find_nWire_Pin:: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'20190416 top
'20170808 evans.lo
Public Function DC_CreateTestName(TestItem As String)
    On Error GoTo ErrorHandler
    
    Dim tmpTName() As String
    Dim Index As Long
    Dim TestItemName As String: TestItemName = TheExec.DataManager.InstanceName
    gDCArrTestName = Split(gDCTestNameTemplate, "_")
    
'    If TestItem Like "*_*" Then
'        tmpTname = Split(TestItem, "_")
'        gDCTestNameIndex = UBound(tmpTname) + 1
'        For Index = 0 To UBound(tmpTname)
'            gDCArrTestName(Index) = tmpTname(Index)
'        Next Index
'    Else
'        gDCTestNameIndex = 1
    gDCArrTestName(DC_TNAME_TESTITEM) = TestItem
    
    If TestItemName Like "*LEAKAGE*" Then
        If TestItemName Like "*_HV*" Then
            gDCArrTestName(8) = "HV"
        ElseIf TestItemName Like "*_LV*" Then
            gDCArrTestName(8) = "LV"
        Else
            gDCArrTestName(8) = "NV"
        End If
    End If
'    End If
    
    Exit Function
ErrorHandler:
    LIB_ErrorDescription ("DC_CreateTestName")
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function DC_GetTestName() As String
    
    On Error GoTo ErrorHandler
    
    DC_GetTestName = Join(gDCArrTestName, "_")
    
    Exit Function
ErrorHandler:
    LIB_ErrorDescription ("DC_GetTestName")
    If AbortTest Then Exit Function Else Resume Next
    
End Function
Public Sub LIB_ErrorDescription(errString As String)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "LIB_ErrorDescription"

    If TheExec.RunMode = runModeDebug Then
       TheExec.Datalog.WriteComment " TheExec.ErrorLogMessage (errString) :  " & errString
    End If
    
Exit Sub
ErrHandler:
     RunTimeError funcName
     If AbortTest Then Exit Sub Else Resume Next
End Sub
'20171030 evans: register dump and compare lib => used for export to file
Private Function ComposeExportString(Site As Variant, ExportRegStatus As REG_STATUS_EXPORT, Index As Long) As String
    Dim ExportDataArray(4) As String
    
    On Error GoTo ErrHandler
    
    ExportDataArray(0) = CStr(Site)
    ExportDataArray(1) = ExportRegStatus.RegName(Index)
    ExportDataArray(2) = ExportRegStatus.RegAddr(Index)
    ExportDataArray(3) = ExportRegStatus.BefData(Index)
    ComposeExportString = Join(ExportDataArray, ",")
    
    Exit Function
ErrHandler:
    LIB_ErrorDescription ("ComposeExportString")
    If AbortTest Then Exit Function Else Resume Next
    
End Function
Private Function CompareExportString(ReadExportString As String, RegName As String) As Boolean

    On Error GoTo ErrHandler
    If InStr(ReadExportString, RegName) <> 0 Then
        CompareExportString = True
    Else
        CompareExportString = False
    End If
    
    Exit Function
ErrHandler:
    LIB_ErrorDescription ("CompareExportString")
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Private Function ComposeRegData(ReadFileString As String, AfterData As String) As String
    
    On Error GoTo ErrHandler
    Dim TmpArray() As String
    
    ComposeRegData = ""
    TmpArray = Split(ReadFileString, ",")
    If TmpArray(S_REG_BEFORE - 1) <> AfterData Then
        ComposeRegData = "," + DiffCheck
    End If
    Exit Function
ErrHandler:
    LIB_ErrorDescription ("ComposeRegData")
    If AbortTest Then Exit Function Else Resume Next
    
End Function

'Private Sub ReSortReadRegData(ReadExportFileArray() As String, FilegReadBySite() As REG_FILE_READ)
'
'    On Error GoTo ErrHandler
'
'    Dim SiteObject As Object
'    Dim Index As Long
'    Dim Count As Long
'    Dim iSite As Variant
'
'    If gbRegDumpOfflineCheck = True Then
'        Set SiteObject = TheExec.Sites.Existing
'    Else
'        Set SiteObject = TheExec.Sites.Selected
'    End If
'
'    Count = 0
'
'    For Each iSite In SiteObject
'        For Index = 0 To UBound(ReadExportFileArray)
'            If InStr(ReadExportFileArray(Index), CStr(iSite)) = 1 Then
'                ReDim Preserve FilegReadBySite(iSite).READDATA(Count)
'                FilegReadBySite(iSite).READDATA(Count) = ReadExportFileArray(Index)
'                Count = Count + 1
'            End If
'        Next Index
'        Count = 0
'    Next iSite
'
'    Exit Sub
'
'ErrHandler:
'    LIB_ErrorDescription ("ReSortReadRegData")
'    If AbortTest Then Exit Sub Else Resume Next
'End Sub
''20190416 end

Public Function Get_nWire_Name(NWire As Variant, port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_PowerSequence_pa As String) ''Support multiple nWire port 20170718
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Get_nWire_Name"
' Nwire is input and can be port name or pin name
' Will output relative name for port/ac_spec/pin/global_spec for powerup_seq
    'Eg. nWire = "XI0_Port"
    'XI0_Port,XI0_Freq_VAR, XI0_PA
    'XI0_Diff_Port,XI0_Diff_Freq_VAR, XI0_Diff_PA
    Dim remove_name As String, key_name As String
    NWire = LCase(NWire)
    If NWire Like "*_port" Then
        remove_name = "_port"
    ElseIf NWire Like "*_freq_var" Then
        remove_name = "_freq_var"
    ElseIf NWire Like "*_pa" Then
        remove_name = "_pa"
    Else
'        TheExec.ErrorLogMessage NWire & "is Wrong nWire name (should be as A_port, A_Freq_Var or A_PA)"
        remove_name = "" 'if it is XI0/RT_CLK32768/CLK_IN ...
    End If
    key_name = UCase(Replace(NWire, remove_name, ""))
    port_pa = key_name & "_PORT"
    ac_spec_pa = key_name & "_FREQ_VAR"
    global_spec_PowerSequence_pa = key_name & "_Port_PowerSequence_GLB"
    pin_pa = key_name & "_PA"

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Disable_FRC(nWire_ports As String, Optional DisConnectFRC As Boolean = False) ''Support multiple nWire port 20170718
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Disable_FRC"

' nWire_ports  can be port name or pin name
' If it is blank, will assume to use all nWire ports
    'Eg. nWire_ports = "XI0_Port, RT_CLK32768_Port, XIN_Port"
    Dim nWire_port_ary() As String
    Dim nwp As Variant, all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim Site As Variant
    If nWire_ports = "" Then nWire_ports = nWire_Ports_GLB
    nWire_port_ary = Split(nWire_ports, ",")
    ' Convert nWire_ports to all_ports and all_pins
    For Each nwp In nWire_port_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        If all_ports = "" Then
            all_ports = port_pa
            all_pins = pin_pa
        Else
            all_ports = all_ports & "," & port_pa
            all_pins = all_pins & "," & pin_pa
        End If
    Next nwp
    
    TheExec.Datalog.WriteComment "******************  Disable freerunning clock " & all_ports & " ****************"
    For Each Site In TheExec.Sites
        TheHdw.Protocol.ports(all_ports).Halt
        TheHdw.Protocol.ports(all_ports).Enabled = False
    Next Site
    
    If DisConnectFRC = True Then
        TheExec.Datalog.WriteComment "******************  Disconnect nWire pins " & all_pins & " ****************"
        TheHdw.Digital.Pins(all_pins).Disconnect
    End If

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Enable_FRC(nWires As String, Optional ConnectFRC As Boolean = False) ''Support multiple nWire port 20170718
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Enable_FRC"
' nWires  can be port name or pin name
' If it is blank, will assume to use all nWire ports
    'Eg. nWires = "XI0_Port, RT_CLK32768_Port, XIN_Port"
    Dim nWires_ary() As String
    Dim nwp As Variant, all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim PLL_Lock As New SiteLong
    Dim port_level_value As Double
    Dim FreeRunFreq As Double
    Dim Site As Variant

    If nWires = "" Then nWires = nWire_Ports_GLB
    nWires_ary = Split(nWires, ",")
    ' Convert nWires to all_ports and all_pins
    For Each nwp In nWires_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        If all_ports = "" Then
            all_ports = port_pa
            all_pins = pin_pa
        Else
            all_ports = all_ports & "," & port_pa
            all_pins = all_pins & "," & pin_pa
        End If
    Next nwp
    
    If ConnectFRC = True Then
        TheHdw.Digital.Pins(all_pins).Connect
        TheExec.Datalog.WriteComment "Connect nWire pins " & all_pins
    End If
    
    TheHdw.Protocol.ports(all_ports).Enabled = True
    TheHdw.Protocol.ports(all_ports).NWire.ResetPLL
    TheHdw.Wait 0.001
    Call TheHdw.Protocol.ports(all_ports).NWire.Frames("RunFreeClock").Execute
    TheHdw.Protocol.ports(all_ports).IdleWait
    TheExec.Datalog.WriteComment "Enable nWire Clock " & all_ports
    
    '****print out to data log about nWire clock condition
    nWires_ary = Split(all_ports, ",")
    For Each nwp In nWires_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        
        For Each Site In TheExec.Sites
            If TheHdw.Protocol.ports(nwp).NWire.IsPLLLocked = False Then
                PLL_Lock = 0
            Else
                PLL_Lock = 1
            End If
        Next Site
        
        FreeRunFreq = 1 / TheHdw.Digital.Timing.Period(nwp) / 1000000
        If TheExec.TesterMode = testModeOffline Then
            For Each Site In TheExec.Sites.Selected
                PLL_Lock = 1
                FreeRunFreq = TheExec.Specs.ac.Item(ac_spec_pa).CurrentValue / 1000000  'offline
             Next Site
        End If
    
        
        TheExec.Flow.TestLimit PLL_Lock, 1, 1, tlSignGreaterEqual, tlSignLessEqual, TName:="nWire " & nwp & " PLL_Lock" 'BurstResult=1:Pass
        If LCase(nwp) Like "*diff*" Then
            port_level_value = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.Value(chVid)
            TheExec.Datalog.WriteComment "********** freerunning clock(" & nwp & ") = " & Format(FreeRunFreq, "0.000") & " Mhz, Vid = " & port_level_value
        Else
            port_level_value = TheHdw.Digital.Pins(pin_pa).Levels.Value(chVih)
            TheExec.Datalog.WriteComment "********** freerunning clock(" & nwp & ") = " & Format(FreeRunFreq, "0.000") & " Mhz, Vih = " & port_level_value
        End If
    Next nwp
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'Public Function Disconnect_FRC(nWires As String) ''Support multiple nWire port 20170718
'On Error GoTo ErrHandler
'Dim funcName As String:: funcName = "Disconnect_FRC"
'' nWires  can be port name or pin name
'' If it is blank, will assume to use all nWire ports
'    'Eg. nWires = "XI0_Port, RT_CLK32768_Port, XIN_Port"
'    Dim nWires_ary() As String
'    Dim nwp As Variant, all_ports As String, all_pins As String
'    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
'
'    If nWires = "" Then nWires = nWire_Ports_GLB
'    nWires_ary = Split(nWires, ",")
'    ' Convert nWires to all_ports and all_pins
'    For Each nwp In nWires_ary
'        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
'        If all_ports = "" Then
'            all_ports = port_pa
'            all_pins = pin_pa
'        Else
'            all_ports = all_ports & "," & port_pa
'            all_pins = all_pins & "," & pin_pa
'        End If
'    Next nwp
'
'    TheExec.Datalog.WriteComment "Disconnect nWire pins " & all_pins
'    TheHdw.Digital.Pins(all_pins).Disconnect
'
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function Connect_FRC(nWires As String) ''Support multiple nWire port 20170718
'On Error GoTo ErrHandler
'Dim funcName As String:: funcName = "Connect_FRC"
'' nWires  can be port name or pin name
'' If it is blank, will assume to use all nWire ports
'    'Eg. nWires = "XI0_Port, RT_CLK32768_Port, XIN_Port"
'    Dim nWires_ary() As String
'    Dim nwp As Variant, all_ports As String, all_pins As String
'    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
'
'    If nWires = "" Then nWires = nWire_Ports_GLB
'    nWires_ary = Split(nWires, ",")
'    ' Convert nWires to all_ports and all_pins
'    For Each nwp In nWires_ary
'        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
'        If all_ports = "" Then
'            all_ports = port_pa
'            all_pins = pin_pa
'        Else
'            all_ports = all_ports & "," & port_pa
'            all_pins = all_pins & "," & pin_pa
'        End If
'    Next nwp
'
'    TheExec.Datalog.WriteComment "Connect nWire pins " & all_pins
'    TheHdw.Digital.Pins(all_pins).Connect
'
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

Public Function Meas_FRC(nWire_ports As String)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Meas_FRC"

    'Eg. nWire_ports = "XI0_Port, RT_CLK32768_Port, XIN_Port"
    Dim nWire_port_ary() As String
    Dim nwp As Variant, meas_freq As New PinListData, Site As Variant
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim PA_Vicm As Double, PA_Vid As Double, PA_Vihd As Double, PA_Vild As Double
    Dim pinlist_pa As New PinList
    
    If nWire_ports = "" Then nWire_ports = nWire_Ports_GLB
    nWire_port_ary = Split(nWire_ports, ",")
    For Each nwp In nWire_port_ary
        Get_nWire_Name CStr(nwp), port_pa, ac_spec_pa, pin_pa, global_spec_pa
        If port_pa Like "*DIFF*" Then
            TheHdw.Digital.Pins(pin_pa).DifferentialLevels.Value(chVod) = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.Value(chVid) / 4
        Else
            TheHdw.Digital.Pins(pin_pa).Levels.Value(chVoh) = TheHdw.Digital.Pins(pin_pa).Levels.Value(chVih) / 4
        End If
        pinlist_pa.Value = pin_pa
        Freq_MeasFreqSetup pinlist_pa, 0.001
        Freq_MeasFreqStart pinlist_pa, 0.001, meas_freq
            
        If TheExec.TesterMode = testModeOffline Then
            For Each Site In TheExec.Sites
                meas_freq.Pins(0).Value = TheExec.Specs.ac(ac_spec_pa).CurrentValue
            Next Site
        End If
                        
        For Each Site In TheExec.Sites
            If port_pa Like "*DIFF*" Then
                PA_Vicm = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.Value(chVicm)
                PA_Vid = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.Value(chVid)
                PA_Vihd = PA_Vicm + PA_Vid / 2
                PA_Vild = PA_Vicm - PA_Vid / 2
                TheExec.Datalog.WriteComment "  FreeRunFreq (" & pin_pa & ") : " & Format(meas_freq.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(PA_Vihd, "0.000") & " v , clock_Vil: " & Format(PA_Vild, "0.000") & " v"
            Else
                TheExec.Datalog.WriteComment "  FreeRunFreq (" & pin_pa & ") : " & Format(meas_freq.Pins(pin_pa).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(pin_pa).Levels.Value(chVih), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(pin_pa).Levels.Value(chVil), "0.000") & " v"
            End If

        Next Site
    Next nwp
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function StartProfile(PinName As String, WhatToCapture As String, SampleRate As Double, SampleSize As Double, CapSignalName As String, Slot_Type As String, _
Optional Meter_I_Range As Double, Optional Meter_V_Range As Double)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "StartProfile"

    ' Wait if another capture is running
    Do While TheHdw.DCVS.Pins(PinName).Capture.IsRunning = True
    Loop
    
    ' Clear capture memory
    TheHdw.DCVS.Pins(PinName).ClearCaptureMemory
    
    ' Create a SIGNAL to set up instrument
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Add CapSignalName

    ' Set this as the default signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.DefaultSignal = CapSignalName

    ' Define the signal used for the capture
    With TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName)
        .Reinitialize
        If (UCase(WhatToCapture) = "I") Then
            .Mode = tlDCVSMeterCurrent
            '.Range = Meter_I_Range
        Else
            .Mode = tlDCVSMeterVoltage
            '.Range = Meter_V_Range
        End If
        .SampleRate = SampleRate
        .SampleSize = SampleSize

    End With

    ' Setup the hardware by loading the signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName).LoadSettings

    ' Start the capture
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName).Trigger

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SplitPinByinstrument(PinName As String, ByRef HexPins As String, ByRef UVSPins As String)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "SplitPinByinstrument"

Dim Profile_AllPin() As String
Dim PinCnt As Long
Dim Pin As Variant
    TheExec.DataManager.DecomposePinList PinName, Profile_AllPin, PinCnt
        
        For Each Pin In Profile_AllPin
            Dim SlotType As String
            SlotType = GetInstrument(CStr(Pin), 0)
            If (LCase(SlotType) = "hexvs") Then
            
                If HexPins = "" Then
                    HexPins = Pin
                Else
                    HexPins = HexPins + "," + Pin
                End If
            ElseIf (LCase(SlotType) = "vhdvs") Then
                
                If UVSPins = "" Then
                    UVSPins = Pin
                Else
                    UVSPins = UVSPins + "," + Pin
                End If
            End If
        Next Pin

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ProfileAutoResolution(SlotType As String, measuretime As Double, ByRef SampleSize As Double, ByRef SampleRate As Double)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "ProfileAutoResolution"

'**************************HexVs***********************************
Dim HexVsMaxSampleSize As Double: HexVsMaxSampleSize = 256000#
Dim HexVsMaxSampleRate As Double: HexVsMaxSampleRate = 25000000#
Dim HexMaxtime As Double: HexMaxtime = HexVsMaxSampleSize * 1024 / HexVsMaxSampleRate
'******************************************************************

'****************************UVS***********************************
Dim UVSMaxSampleSize As Double: UVSMaxSampleSize = 16384
Dim UVSMaxSampleRate As Double: UVSMaxSampleRate = 200000#
Dim UVSMaxtime As Double: UVSMaxtime = UVSMaxSampleSize * (2 ^ 12) / UVSMaxSampleRate
'******************************************************************

Dim RealRate As Double
Dim i As Integer
Dim prediff As Double
Dim posdiff As Double
''Time * SampleRate = SampleSize
    Select Case SlotType
        Case "HEX"
            If measuretime > HexMaxtime Then
                SampleSize = HexVsMaxSampleSize
                SampleRate = HexVsMaxSampleRate / 1024
            Else
                RealRate = HexVsMaxSampleSize / measuretime
                For i = 1 To 1024
                    If RealRate > (HexVsMaxSampleRate / i) Then
                        If i = 1 Then
                            SampleRate = (HexVsMaxSampleRate / i)
                            SampleSize = SampleRate * measuretime
                            Exit For
                        Else
                            prediff = Abs((HexVsMaxSampleRate / (i - 1)) - RealRate)
                            posdiff = Abs((HexVsMaxSampleRate / i) - RealRate)
                            If prediff > posdiff Then
                                SampleRate = (HexVsMaxSampleRate / i)
                                SampleSize = SampleRate * measuretime
                            Else
                                SampleRate = (HexVsMaxSampleRate / (i - 1))
                                SampleSize = SampleRate * measuretime
                            End If
                            Exit For
                        End If
                    End If
                Next i
            End If
        Case "UVS"
            If measuretime > UVSMaxtime Then
                SampleSize = UVSMaxSampleSize
                SampleRate = UVSMaxSampleRate / (2 ^ 12)
            Else
                RealRate = UVSMaxSampleSize / measuretime
                For i = 0 To 12
                    If RealRate > (UVSMaxSampleRate / (2 ^ i)) Then
                        If i = 1 Then
                            SampleRate = (UVSMaxSampleRate / (2 ^ i))
                            SampleSize = SampleRate * measuretime
                            Exit For
                        Else
                            prediff = Abs((UVSMaxSampleRate / 2 ^ (i - 1)) - RealRate)
                            posdiff = Abs((UVSMaxSampleRate / (2 ^ i)) - RealRate)
                            If prediff > posdiff Then
                                SampleRate = (UVSMaxSampleRate / (2 ^ i))
                                SampleSize = SampleRate * measuretime
                            Else
                                SampleRate = (UVSMaxSampleRate / 2 ^ (i - 1))
                                SampleSize = SampleRate * measuretime
                            End If
                            Exit For
                        End If
                    End If
                Next i
            End If
    End Select
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Function Bin2Dec_rev_Fractional(sMyBin As String) As Variant
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Bin2Dec_rev_Fractional"

    Dim x As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        Bin2Dec_rev_Fractional = Bin2Dec_rev_Fractional + Mid(sMyBin, iLen - x + 1, 1) * 2 ^ (-x - 1)
    Next
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GlobalRegMap_Initialize() As Boolean

    On Error GoTo ErrHandler
    
    Dim lastRowIndex As Range
    Dim LastRow As Double
    Dim RegNameIndex As Integer
    Dim CheckRegName As String
    Dim FieldWidth As Long
    Dim Index As Integer
    
    If gbGlobalAddrMap = False Then
        Dim Row As Double
        Dim GlobalAddressMapSheet As Object
        
    'Find the Last Row Index of the GlobalAddressMap
    '-------------------------------------------------------------
        
        'Worksheets("GlobalAddressMap").Activate
        Set GlobalAddressMapSheet = ThisWorkbook.Sheets("AHB_register_map")
        Set lastRowIndex = GlobalAddressMapSheet.Range("A65536").End(xlUp)
        LastRow = lastRowIndex.Row
    '-------------------------------------------------------------
'20180503 evans.lo : For AHB address
        RegNameIndex = 0
        ReDim glGlobalAddrMap(RegNameIndex)
        ReDim glDictGlobalAddrMap(RegNameIndex)
        CheckRegName = GlobalAddressMapSheet.Cells(GLOBAL_ADDR_MAP_INDEX.G_START_ROW, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value
'20180503 evans.lo : For AHB Field Mask
        ReDim Preserve gsAHBFieldName(LastRow - GLOBAL_ADDR_MAP_INDEX.G_START_ROW)
        ReDim Preserve glAHBFieldMask(LastRow - GLOBAL_ADDR_MAP_INDEX.G_START_ROW)
        
        For Row = GLOBAL_ADDR_MAP_INDEX.G_START_ROW To LastRow
'20180503 evans.lo : For AHB address
            If CheckRegName <> GlobalAddressMapSheet.Cells(Row + 1, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value Then
                glGlobalAddrMap(RegNameIndex) = CLng(Replace(UCase(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_ADDR).Value), "0X", "&H"))
                glDictGlobalAddrMap(RegNameIndex) = CheckRegName 'GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value
                CheckRegName = GlobalAddressMapSheet.Cells(Row + 1, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME).Value
                RegNameIndex = RegNameIndex + 1
                If Len(CheckRegName) > 0 Then
                    ReDim Preserve glGlobalAddrMap(RegNameIndex)
                    ReDim Preserve glDictGlobalAddrMap(RegNameIndex)
                End If
            End If
'20180503 evans.lo : For AHB Field Mask
            gsAHBFieldName(Row - GLOBAL_ADDR_MAP_INDEX.G_START_ROW) = UCase(Trim(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_NAME))) & "_" & UCase(Trim(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_FIELD).Value)) '20190915
            FieldWidth = 0
            For Index = 0 To CLng(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REG_FIELD_Width)) - 1
                FieldWidth = FieldWidth + 2 ^ Index
            Next Index
            FieldWidth = FieldWidth * 2 ^ CLng(GlobalAddressMapSheet.Cells(Row, GLOBAL_ADDR_MAP_INDEX.G_REF_FIELD_Offset))
            glAHBFieldMask(Row - GLOBAL_ADDR_MAP_INDEX.G_START_ROW) = CLng("&H" & Mid(CStr(Hex(Not FieldWidth)), 7, 2))
        Next Row
        gbGlobalAddrMap = True
    End If
    
    GlobalRegMap_Initialize = gbGlobalAddrMap
    
Exit Function
    
ErrHandler:
    LIB_ErrorDescription ("GlobalRegMap_Initialize")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20180503 evans : Get AHB Field Mask
Public Function GetAHBFieldMask(FieldName As String) As Long

    On Error GoTo ErrHandler
    Dim Index As Integer
    
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
    If AbortTest Then Exit Function Else Resume Next
End Function





'Public Function AHB_WRITEDSC_Trim(Address As Long, Data As SiteLong, BitOffset As Long, OffsetVal As SiteLong) As Long
'
'Dim dummypat As New PatternSet  'CT 081417
'
'Dim Addr As New SiteLong
'Addr = ToSiteLong(Address)
'
'On Error GoTo ErrHandler
'
'If TheExec.Sites.Selected.Count = 0 Then
'   TheExec.Datalog.WriteComment "*** SK20170926 If any site selected, bypass the AHB_WRITEDSC **************************************************"
'   Exit Function
'End If
'
'
'dummypat = ".\patterns\CPU\JTAG\DD_SILA0_C_FULP_JT_XXXX_WIR_JTG_UNS_ALLFV_SI_TC2AHB_WRDSC_1_A0_1710171804.PAT"
'
'
'Data = Data.BitwiseAnd(2 ^ (BitOffset) - 1).BitwiseOr(OffsetVal.ShiftLeft(BitOffset))
'
'Write_24bits dummypat, "GPIO26", Data, Addr
'
'
'Exit Function
'
'
'ErrHandler:
'    If AbortTest Then Exit Function Else Resume Next
'
'
'End Function

'20171120 mask for trim dsc test function

'Public Sub testmask()
'On Error GoTo ErrHandler
'Dim funcName As String:: funcName = "testmask"
'
'    Dim Data As New SiteLong
'    Dim mask As New SiteLong
'    Data = &H7
'    mask = &H7
''For example : DATA = [7:4], MASK = [3:1]
'    AHB_WRITEDSC_TrimWithMask 0, Data, mask, 4, 1
'
'Exit Sub
'ErrHandler:
'     RunTimeError funcName
'     If AbortTest Then Exit Sub Else Resume Next
'End Sub

''20171120 evans add mask for trim dsc
'Public Function AHB_WRITEDSC_TrimWithMask(Address As Long, Data As SiteLong, mask As SiteLong, Optional Offset As Long = 0, Optional maskoffset As Long = 0) As Long
'
'Dim dummypat As New PatternSet  'CT 081417
'
'Dim Addr As New SiteLong
'Dim tmpMask As New SiteLong
'Addr = ToSiteLong(Address)
'
'On Error GoTo ErrHandler
'
'If TheExec.Sites.Selected.Count = 0 Then
'   TheExec.Datalog.WriteComment "*** SK20170926 If any site selected, bypass the AHB_WRITEDSC **************************************************"
'   Exit Function
'End If
'
'
'dummypat = ".\patterns\CPU\JTAG\DD_SILA0_C_FULP_JT_XXXX_WIR_JTG_UNS_ALLFV_SI_TC2AHB_WRDSC_1_A0_1710171804.PAT"
'
'tmpMask = mask.ShiftLeft(maskoffset)
'Data = Data.ShiftLeft(Offset).BitwiseOr(tmpMask)
'
'Write_24bits dummypat, "GPIO26", Data, Addr
'
'
'Exit Function
'
'
'ErrHandler:
'    If AbortTest Then Exit Function Else Resume Next
'
'
'End Function



'' ****************************************** Example : how to use AHB_READDSC ************************************************
''
'' Previous Project (Imola or SStone): AHB_READDSC BUCK0_HP2_CFG_0 , regval
''
'' Avus Project : 1. Read Register(Same as previous Project): AHB_READDSC BUCK0_HP2_CFG_0.Addr , regval
''                2. Read Register By Field : AHB_READDSC BUCK0_HP2_CFG_0.Addr , regval, BUCK0_HP2_CFG_0.CFG_CC_BLK_SET
''
''*****************************************************************************************************************************
''2018/05/26
'Public Function AHB_READDSC(Address As Long, Data As SiteLong, Optional Field_Mask As Long = 0, Optional bDBGlog As Boolean = True, Optional dWaitTime As Double = 10 * ms) As SiteLong
'On Error GoTo ErrHandler
'Dim dummypat        As New PatternSet       'CT 081417
'Dim mS_PattArray()  As String, mL_PatCount      As Long
'Dim BitAND          As Long, Offset             As Long
'Dim BitANDStr       As String, CalcData         As New SiteLong
'Dim Site As Variant
'
'
'    If TheExec.TesterMode = testModeOffline Then Exit Function '20180910 Need to disable later
'
'    If TheExec.sites.Selected.Count = 0 Then
'       TheExec.Datalog.WriteComment "*** If no Site alive, bypass the AHB_READDSC **************************************************"
'       Exit Function
'    End If
'    Address = Address And &HFFFF&
'    dummypat.Value = OTP_GetPatListFromPatternSet_PATT0(g_sAHB_READ_PAT, mS_PattArray, mL_PatCount)
'
'
'    Data_read dummypat, g_sTDI, g_sTDO, Address, Data, bDBGlog, dWaitTime
'
'    If Field_Mask > 0 Then
'        BitAND = (&HFF) Xor Field_Mask
'        BitANDStr = auto_Dec2Bin_OTP(BitAND, 8)
'        Offset = InStr(StrReverse(BitANDStr), "1") - 1
'        CalcData = Data.bitwiseand(BitAND).ShiftRight(Offset)
'        Data = CalcData
'        For Each Site In TheExec.sites
'                        'needs data print out in the datalog file.  Please use Top level "False" if debug print is not needed.
'            If bDBGlog = True Then TheExec.Datalog.WriteComment "Address-h'" & Hex(Address) & "(d'" & (Address And &HFF) & ")/" & "Data-" & Hex(Data(Site))
'            If bDBGlog = True Then Debug.Print "Address-h'" & Hex(Address) & "(d'" & (Address And &HFF) & ")/" & "Data-" & Hex(Data(Site))
'        Next Site
'    End If
'
'Exit Function
'
'ErrHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'Public Function Data_read(PatName As PatternSet, JTAG_TDI As String, JTAG_TDO As String, Addr As Long, ByRef Data As SiteLong, Optional bDBGlog As Boolean = True, Optional dWaitTime As Double = 10 * ms)
'
''here we are using digsource and digcap at the same time so we write to the Address and read from the Address at the same time
'
'On Error GoTo ErrHandler
'
'
'    Dim addrwidth As Long
'    Dim DataWidth As Long
'
'    Dim i As Long
'    Dim dataout As New SiteLong
'   Dim Address As New SiteLong
'
'     If TheExec.sites.Active.Count = 0 Then Exit Function
'
'
'    Address = ToSiteLong(Addr)
'
'    addrwidth = 16
'    DataWidth = 16
'
'
'        Dim addressSerial As New DSPWave, AddressWave As New DSPWave
'
'        For Each Site In TheExec.sites
'             addressSerial.CreateConstant Address, 1, DspLong
'             addressSerial = addressSerial.ConvertStreamTo(tldspSerial, addrwidth, 0, Bit0IsMsb)
'             AddressWave = addressSerial.Copy.repeat(2)
'             'AddressWave = addressSerial.Copy.repeat(1) '20171118
'        Next Site
'
'
'        Dim SignalName As String
'        Dim WaveDef As String
'        WaveDef = "Wavedef"
'        SignalName = "OnlyAddress"
'
'                  'thehdw.Digital.Patgen.Halt
'     'TheHdw.Wait 10
'
'      addrwidth = 32
'
'          TheHdw.Patterns(PatName).Load
'
'         For Each Site In TheExec.sites.Selected
'
'           TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & Site, AddressWave, True
'          '  TheExec.WaveDefinitions.CreateWaveDefinition WaveDef, AddressWave, True
'
'                     TheHdw.DSSC.Pins(JTAG_TDI).Pattern(PatName).Source.Signals.Add SignalName
'                     With TheHdw.DSSC(JTAG_TDI).Pattern(PatName).Source.Signals(SignalName)
'                            .Reinitialize
'                            .WaveDefinitionName = WaveDef & Site
'                            .Amplitude = 1
'                            .SampleSize = addrwidth 'dressWave.SampleSize 'addrwidth '20171118
'                            .LoadSamples
'                            .LoadSettings
'
'                     End With
'
'
'         Next Site
'
'
'        TheHdw.DSSC(JTAG_TDI).Pattern(PatName).Source.Signals.DefaultSignal = SignalName
'
'         'setup capture
'          Dim DataArray As New DSPWave
'          DataWidth = 8
'          DataArray.CreateConstant 0, 8
'         Call DSSC_Capture_Setup(PatName, JTAG_TDO, "dataSigAHB", DataWidth, DataArray, dWaitTime)
'
'
''          TheHdw.Patterns(PatName).Start ("")
'
'        'TheHdw.Wait 1 * ms
'
'    ' Bind capture results to DSPWave object
''    dataarray = TheHdw.DSSC.Pins(JTAG_TDO).Pattern(PatName).Capture.Signals(SignalName).DSPWave
'
'         ' Bypass DSP computing, use HOST computer
'          TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
'         ' Halt on opcode to make sure all samples are capture.
'          TheHdw.Digital.Patgen.HaltWait
'
''        Dim Tarray(7) As Long
''        Dim outarray(7) As Long
''        Dim test As New DSPWave
''        Dim TmpVal As New SiteLong
''
''      test.CreateConstant 0, 8, DspDouble
''
''
''
''        test.Element(0) = 1
''        test.Element(1) = 0
''        test.Element(2) = 0
''        test.Element(3) = 0
''        test.Element(4) = 1
''        test.Element(5) = 1
''        test.Element(6) = 0
''        test.Element(7) = 0
''
''       test = test.ConvertDataTypeTo(DspLong)
''
''
'      ' TmpVal = test.Element(7) * 2 ^ 7 + test.Element(6) * 2 ^ 6 + test.Element(5) * 2 ^ 5 + test.Element(4) * 2 ^ 4 + test.Element(3) * 2 ^ 3 + test.Element(2) * 2 ^ 2 + test.Element(1) * 2 ^ 1 + test.Element(0) * 2 ^ 0
''
''        Dim test As New DSPWave
''         test.CreateConstant 0, 8, DspDouble
'
'
''
''        test.Element(0) = 1
''        test.Element(1) = 0
''        test.Element(2) = 0
''        test.Element(3) = 0
''        test.Element(4) = 1
''        test.Element(5) = 1
''        test.Element(6) = 0
''        test.Element(7) = 0
'
'
'
''         test = test.ConvertDataTypeTo(DspLong)
'
'             For Each Site In TheExec.sites
'
'                    If DataArray.SampleSize <> 8 Then
'                      TheExec.Datalog.WriteComment "DataArray.SampleSize=" & CStr(DataArray.SampleSize) '20171118
'                      'Stop
'                    End If
'
'                Data(Site) = (DataArray.Element(7) * 2 ^ 7 + DataArray.Element(6) * 2 ^ 6 + DataArray.Element(5) * 2 ^ 5 + DataArray.Element(4) * 2 ^ 4 + DataArray.Element(3) * 2 ^ 3 + DataArray.Element(2) * 2 ^ 2 + DataArray.Element(1) * 2 ^ 1 + DataArray.Element(0) * 2 ^ 0)
'
'
'
'
'
'
'           Next Site
'
'
'
'            For Each Site In TheExec.sites
'              'Buck6 needs data print out in the datalog file.  Please use Top level "False" if debug print is not needed.
'              If bDBGlog = True Then TheExec.Datalog.WriteComment "Address-h'" & Hex(Addr) & "(d'" & (Addr And &HFF) & ")/" & "Data-" & Hex(Data(Site))
'              If bDBGlog = True Then Debug.Print "Address-h'" & Hex(Addr) & "(d'" & (Addr And &HFF) & ")/" & "Data-" & Hex(Data(Site))
'          Next Site
'
'
'
'
'       Exit Function
'
'
'ErrHandler:
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function

''******************************************************************************
''' Digital Signal Capture utilities
'''******************************************************************************
'Public Function DSSC_Capture_Setup(PatName As PatternSet, DigCapPin As String, _
'                SignalName As String, SampleSize As Long, Capwave As DSPWave, Optional dWaitTime As Double = 10 * ms)
'
'    On Error GoTo ErrHandler
'    'TheHdw.Patterns(PatName).Load
'    With TheHdw.DSSC.Pins(DigCapPin).Pattern(PatName).Capture.Signals
'        .Reinitialize
'        .Add (SignalName)
'        With .Item(SignalName)
'            .Reinitialize
'            .SampleSize = SampleSize
'            .LoadSettings
'        End With
'    End With
'
'
'    ' Bind capture results to DSPWave object
' 'CapWave = TheHdw.DSSC.Pins(DigCapPin).Pattern(PatName).Capture.Signals(SignalName).DSPWave   'WAS.  20171118 REMOVE
'
'    'Bypass DSP computing, use HOST computer '20171118
'    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
'    'halt on opcode to make sure all samples are capture.
'    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
'    ' Bind capture results to DSPWave object
'  'CapWave = TheHdw.DSSC.Pins(DigCapPin).Pattern(PatName).Capture.Signals(SignalName).DSPWave
'    TheHdw.Wait dWaitTime ' was fixed 10 * ms 'add wait time for acore RPoly open kelvin alarm
'
'    ''TheHdw.Patterns(PatName).start ("")
'    Call TheHdw.Patterns(PatName).test(pfNever, 0)
'    ' TheHdw.Wait 2
'
'
'          TheHdw.Digital.Patgen.HaltWait
'
'
'
'      ' Bind capture results to DSPWave object
'    Capwave = TheHdw.DSSC.Pins(DigCapPin).Pattern(PatName).Capture.Signals(SignalName).DSPWave   '20171118
'
'
'    Exit Function
'
'ErrHandler:
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function

Public Function Freq_MeasFreqSetup(Pin As PinList, Interval As Double, Optional MeasF_EventSource As FreqCtrEventSrcSel = 1)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Freq_MeasFreqSetup"

    With TheHdw.Digital.Pins(Pin).FreqCtr
        .EventSource = MeasF_EventSource '' VOH
        .EventSlope = Positive
        .Interval = Interval
        .Enable = IntervalEnable
        .Clear
    End With
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Freq_MeasFreqStart(Pin As PinList, Interval As Double, freq As PinListData)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "Freq_MeasFreqStart"


    Dim CounterValue As New PinListData
    Dim Site As Variant
    Dim Result As New SiteLong
    TheHdw.Digital.Pins(Pin).FreqCtr.Clear
    TheHdw.Digital.Pins(Pin).FreqCtr.Start
    
    For Each Site In TheExec.Sites
        CounterValue = TheHdw.Digital.Pins(Pin).FreqCtr.Read
        freq = CounterValue.Math.Divide(Interval)
    Next Site
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetInstrument(PinList As String, Site As Variant) As String
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "GetInstrument"

    Dim chanString As String
    Dim PinName() As String
    Dim NumberPins As Long
    Call TheExec.DataManager.DecomposePinList(PinList, PinName(), NumberPins)
    Call TheExec.DataManager.GetChannelStringFromPinAndSite(PinName(0), Site, chanString)
    Dim slotstr() As String
    Dim slot As Long
    If chanString = "" Then
        MsgBox ("Please check pin type of  " & PinList & " in channel map")
    Else
        slotstr = Split(chanString, ".")
        slot = CLng(slotstr(0))
        GetInstrument = TheHdw.config.Slots(slot).Type
    End If
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ActivateAllSheet()
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "GetInstrument"

  Dim ws As Worksheet
   For Each ws In ActiveWorkbook.Worksheets
   ws.Activate
   Next
   
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'20190416 end

'20171027 evans: exort reg data to worksheet or csv file
'20170814 evans: for CT request
'Public Function GetDataByRegName(start_reg As String, end_reg As String, RegCheck As REG_DATA, Optional FileName As String = "REGCHECK.csv", _
'                                                Optional ExportToSheet As Boolean = True, Optional ExportToFile As Boolean = True)
'
'
'    Dim SiteObject As Object
'    Dim ReadExportFileArray() As String
'
'    On Error GoTo ErrHandler
'
'    If gbRegDumpOfflineCheck = True Then
'        Set SiteObject = TheExec.sites.Existing
'    Else
'        Set SiteObject = TheExec.sites.Selected
'    End If
'
'    If ExportToSheet = False And ExportToFile = False Then
'        TheExec.Datalog.WriteComment "<WARNING> : The alternatives are ExportToSheet and ExportToFile!"
'        Exit Function
'    End If
'
'    GlobalRegMap_Initialize
'
'    If ExportToFile Then
'        Dim mS_HEADER As String
'        Dim mS_FILETYPE As String
'        Dim mS_File  As String
'        Dim mS_FileName As String
'        Dim fs As New FileSystemObject
'        Dim St_ReadTxtFile As TextStream
'        Dim read_count As Long
'
'        mS_FILETYPE = ".csv"
'
'        mS_FileName = "REGCHECK" + mS_FILETYPE 'Join(mS_FileNameArray, "_") + mS_FILETYPE
'        If InStr(mS_FileName, FileName) = 0 And Len(FileName) > 0 Then
'            mS_FileName = FileName
'            If InStr(FileName, mS_FILETYPE) = 0 Then
'                mS_FileName = mS_FileName + mS_FILETYPE
'            End If
'        End If
'
'        mS_File = gS_REGCHECKFileDir + mS_FileName
'        mS_HEADER = "Site,Reg Name,Reg Addr,Before,After"
'
'        Call File_CheckAndCreateFolder(gS_REGCHECKFileDir)
'        If RegCheck = REG_DATA_BEFORE And fs.FileExists(mS_File) Then
'            Call File_CreateAFile(mS_File, mS_HEADER)
'        End If
'
'        read_count = 0
'
'        If RegCheck = REG_DATA_AFTER Then
'
'            ReDim FilegReadBySite(TheExec.sites.Existing.Count - 1) 'offline check
'
'            Set St_ReadTxtFile = fs.OpenTextFile(mS_File, ForReading, True)
'            St_ReadTxtFile.ReadLine 'filter header
'            Do While Not St_ReadTxtFile.AtEndOfStream
'                ReDim Preserve ReadExportFileArray(read_count)
'                ReadExportFileArray(read_count) = St_ReadTxtFile.ReadLine
'                read_count = read_count + 1
'            Loop
'            St_ReadTxtFile.Close
'            Call ReSortReadRegData(ReadExportFileArray, FilegReadBySite)
'        End If
'    End If
'
'    If gbGlobalAddrMap = True Then
'        Dim Index As Long
'        Dim RegData As New SiteLong
'        Dim iSite As Variant
'        Dim RegStatusCheckSheet As Object
'        Dim Count As Long
'        Dim start_index, end_index As Long
'        Dim start_addr As Long, end_addr As Long
'
'        start_addr = -1
'        end_addr = -1
'
'        For Index = 0 To UBound(glDictGlobalAddrMap)
'            If glDictGlobalAddrMap(Index) = start_reg Then start_addr = glGlobalAddrMap(Index)
'            If glDictGlobalAddrMap(Index) = end_reg Then end_addr = glGlobalAddrMap(Index)
'        Next Index
'
'        If start_addr = -1 Or end_addr = -1 Then
'            TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Register Name(it should be from GlobalAddressMap worksheet!)"
'            Exit Function
'        End If
'
'        If end_addr < start_addr Then
'            TheExec.Datalog.WriteComment "<ERROR> : The end register address should larger than start register address"
'            Exit Function
'        End If
'
'        start_index = -1
'        end_index = -1
'
'        For Index = 0 To UBound(glGlobalAddrMap)
'            If glGlobalAddrMap(Index) = start_addr Then start_index = Index
'            If glGlobalAddrMap(Index) = end_addr Then end_index = Index
'        Next Index
'
'        If start_index = -1 Or end_index = -1 Then
'            TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Address(it should be from GlobalAddressMap worksheet!)"
'            Exit Function
'        End If
'
'        If (end_index - start_index) > &HFF Then
'            TheExec.Datalog.WriteComment "<WARNING> : Dump over 100 registers only export to CSV file!"
'            ExportToSheet = False
'            ExportToFile = True
'        End If
'
'        Worksheets("REG_STATUS_CHECK").Activate
'        Set RegStatusCheckSheet = ThisWorkbook.Sheets("REG_STATUS_CHECK")
'
'        If RegCheck = REG_DATA_BEFORE Then
'            If Len(RegStatusCheckSheet.Cells(2, S_REG_BEFORE).Value) > 0 Then
'                RegStatusCheckSheet.UsedRange.ClearContents
'            End If
'        End If
'
'
'        ReDim ExportRegStatusBySite(TheExec.sites.Existing.Count - 1) 'offline check
'        Count = 0
'
'        For Index = 0 To UBound(glGlobalAddrMap)
'            If glGlobalAddrMap(Index) >= start_addr And glGlobalAddrMap(Index) <= end_addr Then
'                Call AHB_READDSC(glGlobalAddrMap(Index), RegData)
'                TheHdw.Wait 0.05
'    '''                TheExec.Datalog.WriteComment "RegAddr:" & glDictGlobalAddrMap(Index)
'                For Each iSite In SiteObject
'                    ReDim Preserve ExportRegStatusBySite(iSite).RegName(Count)
'                    ExportRegStatusBySite(iSite).RegName(Count) = glDictGlobalAddrMap(Index)
'
'                    ReDim Preserve ExportRegStatusBySite(iSite).RegAddr(Count)
'                    ExportRegStatusBySite(iSite).RegAddr(Count) = "0x" & CStr(Hex(glGlobalAddrMap(Index)))
'
'                    If RegCheck = REG_DATA_BEFORE Then
'                        ReDim Preserve ExportRegStatusBySite(iSite).BefData(Count)
'                        ExportRegStatusBySite(iSite).BefData(Count) = CStr(Hex(RegData(iSite)))
'
'                        If ExportToFile = True Then
'                            ReDim Preserve ExportRegStatusBySite(iSite).ExportToFile(Count)
'                            ExportRegStatusBySite(iSite).ExportToFile(Count) = ComposeExportString(iSite, ExportRegStatusBySite(iSite), Count)
'                        End If
'
'                    End If
'
'                    If RegCheck = REG_DATA_AFTER Then
'                        ReDim Preserve ExportRegStatusBySite(iSite).AftData(Count)
'                        ExportRegStatusBySite(iSite).AftData(Count) = CStr(Hex(RegData(iSite)))
'
'                        If ExportToFile = True Then
'                            If CompareExportString(ReadExportFileArray(Count), glDictGlobalAddrMap(Index)) = False Then
'                                TheExec.Datalog.WriteComment "<ERROR> : Please check the Start or End Address(it should be same as Previous Address Setup!)"
'                                Exit Function
'                            Else
'                                ReDim Preserve ExportRegStatusBySite(iSite).ExportToFile(Count)
'                                ExportRegStatusBySite(iSite).ExportToFile(Count) = FilegReadBySite(iSite).READDATA(Count) + ExportRegStatusBySite(iSite).AftData(Count) _
'                                                                                  + ComposeRegData(FilegReadBySite(iSite).READDATA(Count), ExportRegStatusBySite(iSite).AftData(Count))
'                            End If
'                        End If
'
'                    End If
'
'                Next iSite
'                Count = Count + 1
'            End If
'        Next Index
'
'        If ExportToSheet Then
'            Index = 0
'            Count = 2
'            RegStatusCheckSheet.Cells(Index + 1, S_REG_SITE).Value = "Site"
'            RegStatusCheckSheet.Cells(Index + 1, S_REG_NAME).Value = "Reg Name"
'            RegStatusCheckSheet.Cells(Index + 1, S_REG_ADDR).Value = "Reg Addr"
'            RegStatusCheckSheet.Cells(Index + 1, S_REG_BEFORE).Value = "Before"
'            RegStatusCheckSheet.Cells(Index + 1, S_REG_AFTER).Value = "After"
'            For Each iSite In SiteObject
'                For Index = 0 To UBound(ExportRegStatusBySite(iSite).RegName)
'                    RegStatusCheckSheet.Cells(Index + Count, S_REG_SITE).Value = CStr(iSite)
'                    RegStatusCheckSheet.Cells(Index + Count, S_REG_NAME).Value = ExportRegStatusBySite(iSite).RegName(Index)
'                    RegStatusCheckSheet.Cells(Index + Count, S_REG_ADDR).Value = "0x" & ExportRegStatusBySite(iSite).RegAddr(Index)
'                    If RegCheck = REG_DATA_BEFORE Then RegStatusCheckSheet.Cells(Index + Count, S_REG_BEFORE).Value = ExportRegStatusBySite(iSite).BefData(Index)
'                    If RegCheck = REG_DATA_AFTER Then
'                        RegStatusCheckSheet.Cells(Index + Count, S_REG_AFTER).Value = ExportRegStatusBySite(iSite).AftData(Index)
'                        If RegStatusCheckSheet.Cells(Index + Count, S_REG_AFTER).Value <> RegStatusCheckSheet.Cells(Index + Count, S_REG_BEFORE).Value Then
'                            RegStatusCheckSheet.Cells(Index + Count, S_REG_CHECK).Value = DiffCheck
'                        End If
'                    End If
'                Next Index
'                Count = Count + UBound(ExportRegStatusBySite(iSite).RegName) + 1
'            Next iSite
'        End If
'
'        If ExportToFile Then
'            Dim St_WriteTxtFile As TextStream
'
'            Set St_WriteTxtFile = fs.OpenTextFile(mS_File, ForWriting, True)
'            St_WriteTxtFile.WriteLine mS_HEADER
'            For Each iSite In SiteObject
'                For Index = 0 To UBound(ExportRegStatusBySite(iSite).ExportToFile)
'                   St_WriteTxtFile.WriteLine ExportRegStatusBySite(iSite).ExportToFile(Index)
'                Next Index
'            Next iSite
'
'            St_WriteTxtFile.Close
'
'        End If
'
'    End If
'
'    TheExec.Datalog.WriteComment "<GetDataByRegName> : Register dump is completed!"
'
'    Exit Function
'
'ErrHandler:
'    LIB_ErrorDescription ("GetDataByRegName")
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function





Public Function DC_UpdatePinName(PinName As String, Optional TestCond As String, Optional TestTemp As String)
    
    On Error GoTo ErrorHandler
    gDCArrTestName(DC_TNAME_TESTPIN) = Replace(PinName, "_", "-")
    If TestCond <> "" Then
        gDCArrTestName(DC_TNAME_TESTCOND) = TestCond
    End If
    If TestTemp <> "" Then
        gDCArrTestName(DC_TNAME_TESTTEMP) = TestTemp
    End If
'    Dim tmpPinName() As String
'    Dim Index As Long
'    tmpPinName = Split(PinName, "_")
'    For Index = 0 To UBound(tmpPinName)
'        gDCArrTestName(Index + gDCTestNameIndex) = tmpPinName(Index)
'    Next Index
    Exit Function
ErrorHandler:
    LIB_ErrorDescription ("DC_UpdateTestName")
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function PrLoadPattern(PatName As String) As Long
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "PrLoadPattern"

    Dim PattArray() As String
    Dim patt As Variant
    Dim Pat As String
    Dim PatCount As Long
    
    If PatName = "" Then Exit Function
    
    ' Run validation
    Call ValidatePattern(PatName)
    Call PATT_GetPatListFromPatternSet(PatName, PattArray, PatCount)

    For Each patt In PattArray
        Pat = CStr(patt)
        Call ValidatePattern(Pat)
    Next patt

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckIfSheetExists(SheetName As String) As Boolean
On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim m_dlogStr As String

    CheckIfSheetExists = False
    For Each ws In Worksheets
        ''Debug.Print "ws.Name = " & ws.Name
        If LCase(SheetName) = LCase(ws.Name) Then
          CheckIfSheetExists = True
          Exit Function
        End If
    Next ws

    If (CheckIfSheetExists = False) Then
        m_dlogStr = "<WARNING> CheckIfSheetExists:: Sheet " + SheetName + " is NOT existed."
        TheExec.AddOutput m_dlogStr
        TheExec.Datalog.WriteComment m_dlogStr
    End If
    
    Exit Function
ErrorHandler:
    LIB_ErrorDescription ("CheckIfSheetExists")
    If AbortTest Then Exit Function Else Resume Next
End Function
'Move to LIB_Common
Public Function File_CreateAFile(FileName As String, Text As String)

Dim fso As FileSystemObject
Dim fid As TextStream
Dim i As Long
Dim funcName As String: funcName = "File_CreateAFile"

    On Error GoTo ErrHandler
        
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fid = fso.CreateTextFile(FileName, True)
    
    fid.WriteLine (Text)
    fid.Close

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'Public Function File_WriteToFile(FileName As String, Optional WriteLine As String, Optional showPrint As Boolean = False) As Integer
'
'On Error GoTo ErrHandler
'    Dim funcName As String: funcName = "File_WriteToFile"
'    Dim fs As New FileSystemObject
'    Dim St_WriteTxtFile As TextStream
'    Dim j As Long
'    Dim temp_string As String
'
'
'    Set St_WriteTxtFile = fs.OpenTextFile(FileName, ForAppending, True)
'
'    If WriteLine <> "" Then
'
'            St_WriteTxtFile.WriteLine WriteLine
'            St_WriteTxtFile.Close
'
'    Else
'                temp_string = ""
'                If showPrint Then Debug.Print temp_string
'                St_WriteTxtFile.WriteLine temp_string
'        St_WriteTxtFile.Close
'    End If
'
'
'Exit Function
'
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function
'----------------------------------------
'1.Check the Folder Exist Or Not.
'2.Select Y/N for creating the new folder.
'----------------------------------------
Public Function File_CheckAndCreateFolder(Optional Path_FolderName As String = g_sOTPDATA_FILEDIR, Optional CreateFolder As YESNO_type = Yes)
Dim funcName As String: funcName = "CheckAndCreateFolder"
On Error GoTo ErrHandler

Dim mS_TempString As String
Dim mB_DebugPrtDlog As Boolean

mB_DebugPrtDlog = True
mS_TempString = ""

     Dim ffs As New FileSystemObject

     If (Right(Trim(Path_FolderName), 1) <> "\") Then Path_FolderName = Path_FolderName + "\"

     If Not (ffs.FolderExists(Path_FolderName)) Then
        TheExec.Datalog.WriteComment "<Notice!!> The Folder:" + Path_FolderName + " Is Not Exist."
        Select Case CreateFolder
            Case Yes
                ffs.CreateFolder Path_FolderName
                 mS_TempString = "The New Folder:" + Path_FolderName
            Case No
                 mS_TempString = "Skip To Create The New Folder:" + Path_FolderName
        End Select
     Else
                 'mS_TempString = "The Datalog Folder:" + Path_FolderName
     End If

      If (mB_DebugPrtDlog) Then TheExec.Datalog.WriteComment mS_TempString

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function




