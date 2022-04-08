Attribute VB_Name = "LIB_Common"
Option Explicit
'Revision History:
'V0.0 initial bring up
'V0.1 add bintable inital VBT.
'variable declaration
Public Const Version_Lib_Common = "0.1"  'lib version
Public Const DebugPrintEnable = False   'debug print in VBT modules
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

Public Function is_reference_installed(s As String) As Boolean
    Dim X As Variant
    is_reference_installed = False
    For Each X In Application.ActiveWorkbook.VBProject.References
        If s = X.Name Then
            is_reference_installed = True
        End If
    Next X
End Function

Function WorksheetExists(wsName As String, delete As Boolean) As Boolean
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
End Function
Public Function Wait(Time As Double, Optional Debug_Flag As Boolean = False)
'pause few time
    TheHdw.Wait Time
    If Debug_Flag Then
        TheExec.Datalog.WriteComment ("print: Wait time = " + CStr(Time * 1000#) + " mS")
    End If
End Function

Public Sub ErrorDescription(funcName As String)
'error description printing
    Dim TestInstanceName As String
    TestInstanceName = TheExec.DataManager.instanceName
    
    TheExec.Datalog.WriteComment "TestInstance: " & TestInstanceName & ", " & funcName & " error, Err Code: " & err.number & ", Err Description: " & err.Description
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
    
    On Error GoTo errHandler
    
    If original_pin_cnt <> 0 Then
        i = 0   'init
        For Each p In original_ary
            If TheExec.DataManager.ChannelType(p) <> "N/C" Then i = i + 1
        Next p
        
'''''        'redim
'''''        ReDim TempArray(i - 1)
        
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

'''''        For Each p In original_ary
'''''            If TheExec.DataManager.ChannelType(p) <> "N/C" Then
'''''                TempArray(j) = original_ary(j)
'''''                j = j + 1
'''''            Else
'''''                j = j
'''''            End If
'''''        Next p
        
        'return array and pin count
        original_ary = TempArray
        original_pin_cnt = j
    End If
    
    Exit Function
errHandler:
    ErrorDescription ("Trim_NC_Pin")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub RemoveandCopyQualifiers()
    Dim n As Long
    Dim i As Long
    Dim Header As String
    Dim WorkSheetType As String
    Dim Posn As Long
    Dim row As Long
    Dim Opcode As String
    Dim Continue As Boolean
    Dim usedrow As Double

    n = Worksheets.Count

    For i = 1 To n
    Header = Worksheets(i).Cells(1, 1).Value

    Posn = InStr(Header, ",")
    If Posn > 0 Then
    WorkSheetType = Mid(Header, 1, Posn - 1)
    If WorkSheetType = "DTFlowtableSheet" Then
    Worksheets(i).Activate
    usedrow = ActiveSheet.UsedRange.Rows.Count
    ActiveSheet.range(Cells(5, 24), Cells(usedrow, 30)).Select
    ActiveSheet.range(Cells(5, 24), Cells(usedrow, 30)).Copy
    ActiveSheet.range(Cells(5, 44), Cells(usedrow, 50)).Select
    ActiveSheet.range(Cells(5, 44), Cells(usedrow, 50)).PasteSpecial
    ActiveSheet.range(Cells(5, 24), Cells(usedrow, 30)).Select
    ActiveSheet.range(Cells(5, 24), Cells(usedrow, 30)).ClearContents
    End If
    End If

    Next i
End Sub

Public Sub RecoveryandCopyQualifiers()
    Dim n As Long
    Dim i As Long
    Dim Header As String
    Dim WorkSheetType As String
    Dim Posn As Long
    Dim row As Long
    Dim Opcode As String
    Dim Continue As Boolean
    Dim usedrow As Double

    n = Worksheets.Count

    For i = 1 To n
    Header = Worksheets(i).Cells(1, 1).Value

    Posn = InStr(Header, ",")
    If Posn > 0 Then
    WorkSheetType = Mid(Header, 1, Posn - 1)
    If WorkSheetType = "DTFlowtableSheet" Then
    Worksheets(i).Activate
    usedrow = ActiveSheet.UsedRange.Rows.Count
    ActiveSheet.range(Cells(5, 44), Cells(usedrow, 50)).Select
    ActiveSheet.range(Cells(5, 44), Cells(usedrow, 50)).Copy
    ActiveSheet.range(Cells(5, 24), Cells(usedrow, 30)).Select
    ActiveSheet.range(Cells(5, 24), Cells(usedrow, 30)).PasteSpecial
    ActiveSheet.range(Cells(5, 44), Cells(usedrow, 50)).Select
    ActiveSheet.range(Cells(5, 44), Cells(usedrow, 50)).ClearContents
    End If
    End If

    Next i

End Sub

'*****************************************
'******                         TDR ******
'*****************************************
Public Function TDR_Gen_Cal_File() As Long
   ' On Error GoTo errHandler
    Dim ws_chan As Worksheet
    Dim ws_tdr As Worksheet
    Dim wb As Workbook
    Dim chanmap_name As String, tdr_name As String
    Dim row_chan As Long, row_tdr As Long
    Dim site As Variant

    TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

    Set wb = Application.ActiveWorkbook
    chanmap_name = TheExec.CurrentChanMap
    Set ws_chan = wb.Sheets(chanmap_name)
    ws_chan.Cells(3, 6) = "Signal"
    tdr_name = "TDR_DATA_" & chanmap_name
    Call WorksheetExists(tdr_name, True)
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = tdr_name
     Set ws_tdr = wb.Sheets(tdr_name)

    ws_tdr.Cells(1, 1).Value = "Pin Name"
    ws_tdr.Cells(1, 1).Interior.Color = RGB(128, 128, 0)
    For Each site In TheExec.sites.Existing
        ws_tdr.Cells(1, 2).Value = "Chan Site" & site
        ws_tdr.Cells(1, 2).Interior.Color = RGB(128, 128, 0)
        ws_tdr.Cells(1, CLng(site) + 3).Value = "Trace Site " & site
        ws_tdr.Cells(1, CLng(site) + 3).Interior.Color = RGB(128, 128, 0)
    Next site

    row_chan = 7
    row_tdr = 2
    While (ws_chan.Cells(row_chan, 2).Value <> "")
        If ws_chan.Cells(row_chan, 4).Value = "I/O" Then
            For Each site In TheExec.sites.Existing
                    ws_tdr.Cells(row_tdr, 1).Value = ws_chan.Cells(row_chan, 2).Value
                    If ws_chan.Cells(row_chan, 5 + CLng(site)).Value Like "*.ch*" Then
                        ws_tdr.Cells(row_tdr, 2 + 2 * CLng(site)).Value = ws_chan.Cells(row_chan, 5 + CLng(site)).Value
                        ws_tdr.Cells(row_tdr, 3 + 2 * CLng(site)).Value = TheHdw.Digital.Calibration.Channels(ws_chan.Cells(row_chan, 5 + CLng(site)).Value).DIB.trace
                    ElseIf ws_chan.Cells(row_chan, 5 + CLng(site)).Value Like "*site*" Then
                        ws_tdr.Cells(row_tdr, 2 + 2 * CLng(site)).Value = ws_chan.Cells(row_chan, 5 + CLng(0)).Value
                        ws_tdr.Cells(row_tdr, 3 + 2 * CLng(site)).Value = TheHdw.Digital.Calibration.Channels(ws_chan.Cells(row_chan, 5 + CLng(0)).Value).DIB.trace
                    Else
                        MsgBox " select Signal in chanmap sheet  " & chanmap_name
                        Exit Function
                    End If
                    If ws_tdr.Cells(row_tdr, 3 + 2 * CLng(site)).Value > 0.0000000655 Then
                        ws_tdr.Cells(row_tdr, 3 + 2 * CLng(site)).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws_tdr.Cells(row_tdr, 3 + 2 * CLng(site)).Interior.Color = RGB(255, 255, 255)
                    End If

            Next site
            row_tdr = row_tdr + 1
        End If
        row_chan = row_chan + 1
    Wend

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TDR_Read_Compare() As Long
    On Error GoTo errHandler
    Dim ws_chan As Worksheet
    Dim ws_tdr As Worksheet, ws_tdr_cmp As Worksheet
    Dim wb As Workbook
    Dim chanmap_name As String, tdr_name As String, tdr_cmp_name As String
    Dim row_chan As Long, row_tdr As Long, tdr_chan As String
    Dim site As Variant

    TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

    Set wb = Application.ActiveWorkbook
    chanmap_name = TheExec.CurrentChanMap
    Set ws_chan = wb.Sheets(chanmap_name)

    tdr_name = "TDR_DATA_" & chanmap_name
    If WorksheetExists(tdr_name, False) = False Then
        TheExec.AddOutput tdr_name & " does not exist"
        Exit Function
    Else
        Set ws_tdr = wb.Sheets(tdr_name)
    End If

    tdr_cmp_name = "TDR_CMPARE"
    WorksheetExists tdr_cmp_name, True
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = tdr_cmp_name
    Set ws_tdr_cmp = wb.Sheets(tdr_cmp_name)

    ws_tdr_cmp.Cells(1, 1).Value = "Pin Name"
    ws_tdr_cmp.Cells(1, 1).Interior.Color = RGB(128, 128, 0)

    For Each site In TheExec.sites.Existing
        ws_tdr_cmp.Cells(1, 2).Value = "Chan Site" & site
        ws_tdr_cmp.Cells(1, 2).Interior.Color = RGB(128, 128, 0)
        ws_tdr_cmp.Cells(1, CLng(site) + 3).Value = "Org Trace Site " & site
        ws_tdr_cmp.Cells(1, CLng(site) + 3).Interior.Color = RGB(128, 128, 0)
        ws_tdr_cmp.Cells(1, CLng(site) + 4).Value = "Tester Trace Site " & site
        ws_tdr_cmp.Cells(1, CLng(site) + 4).Interior.Color = RGB(128, 128, 0)
    Next site

    row_tdr = 2
    While (ws_tdr.Cells(row_tdr, 2).Value <> "")
        ws_tdr_cmp.Cells(row_tdr, 1).Value = ws_tdr.Cells(row_tdr, 1).Value
        For Each site In TheExec.sites.Existing
                tdr_chan = ws_tdr.Cells(row_tdr, 2 + 2 * CLng(site)).Value
                ws_tdr_cmp.Cells(row_tdr, 2 + 3 * CLng(site)).Value = tdr_chan
                ws_tdr_cmp.Cells(row_tdr, 3 + 3 * CLng(site)).Value = ws_tdr.Cells(row_tdr, 3 + 2 * CLng(site)).Value
                ws_tdr_cmp.Cells(row_tdr, 4 + 3 * CLng(site)).Value = TheHdw.Digital.Calibration.Channels(tdr_chan).DIB.trace
                If Abs(ws_tdr_cmp.Cells(row_tdr, 3 + 3 * CLng(site)).Value - ws_tdr_cmp.Cells(row_tdr, 4 + 3 * CLng(site)).Value) > 0.0000000001 Then
                    ws_tdr_cmp.Cells(row_tdr, 4 + 3 * CLng(site)).Interior.Color = RGB(255, 0, 0)
                Else
                    ws_tdr_cmp.Cells(row_tdr, 4 + 3 * CLng(site)).Interior.Color = RGB(255, 255, 255)
                End If
        Next site
        row_tdr = row_tdr + 1
    Wend

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function TDR_Write() As Long
    Dim ws_chan As Worksheet
    Dim ws_tdr As Worksheet, ws_tdr_cmp As Worksheet
    Dim wb As Workbook
    Dim chanmap_name As String, tdr_name As String, tdr_cmp_name As String
    Dim row_chan As Long, row_tdr As Long, tdr_chan As String
    Dim site As Variant
    Dim tdr_len As Double, tdr_pin As String, row_change_start As Double, row_change_stop As Double
'    On Error GoTo errHandler
    
    TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"
    Set wb = Application.ActiveWorkbook
    chanmap_name = TheExec.CurrentChanMap
    Set ws_chan = wb.Sheets(chanmap_name)
    
    tdr_name = "TDR_DATA_" & chanmap_name
    If WorksheetExists(tdr_name, False) = False Then
        TheExec.AddOutput tdr_name & " does not exist"
        Exit Function
    Else
        Set ws_tdr = wb.Sheets(tdr_name)
    End If
        
    row_tdr = 2
     If LCase(TheExec.CurrentJob) Like "ft*" Then
        row_change_start = 163 - 6 + 1     'refer to  chanmap
        row_change_stop = 188 - 6 + 1      'refer to  chanmap
    ElseIf LCase(TheExec.CurrentJob) Like "cp*" Then
        row_change_start = 163 - 6 + 1      'refer to  chanmap
        row_change_stop = 188 - 5 - 6 + 1   'refer to  chanmap
    End If
    
    While (ws_tdr.Cells(row_tdr, 2).Value <> "")
        For Each site In TheExec.sites.Existing
                tdr_chan = ws_tdr.Cells(row_tdr, 2 + 2 * CLng(site)).Value
                tdr_len = ws_tdr.Cells(row_tdr, 3 + 2 * CLng(site)).Value
                If row_tdr < row_change_start Or row_tdr > row_change_stop Then
'                    theexec.AddOutput "site " & site & "," & row_tdr - 1 & "," & ws_tdr.Cells(row_tdr, 1).Value & "," & _
'                                                            tdr_chan & "," & _
'                                                            ws_tdr.Cells(row_tdr, 3 + 2 * CLng(site)).Value
                    TheHdw.Digital.Calibration.Channels(tdr_chan).DIB.trace = tdr_len
                Else 'overwrite MIPI LP with MIPI HS
'                     theexec.AddOutput "site " & site & "," & row_tdr - 1 & "," & ws_tdr.Cells(row_tdr - 26, 1).Value & "," & _
'                                                            tdr_chan & "," & _
'                                                            ws_tdr.Cells(row_tdr - 26, 3 + 2 * CLng(site)).Value
                    TheHdw.Digital.Calibration.Channels(tdr_chan).DIB.trace = ws_tdr.Cells(row_tdr - 26, 3 + 2 * CLng(site)).Value
               End If
        Next site
        row_tdr = row_tdr + 1
    Wend
    TheExec.Datalog.WriteComment "num of pins = " & row_tdr - 2
    
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function setXY(X As Integer, y As Integer, Optional Device As String) As Long

''''20171103 update
On Error GoTo errHandler
    Dim funcName As String:: funcName = "setXY"
    
    Dim m_chmapName As String
    Dim m_siteCnt As Long
    Dim m_match_flag As Boolean
    
    m_match_flag = False
    m_siteCnt = TheExec.sites.Existing.Count
    m_chmapName = LCase(TheExec.CurrentChanMap)
    
    If (m_siteCnt = 1) Then
        m_match_flag = True
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, X)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)

    ElseIf (m_siteCnt = 2) Then
        If (m_chmapName Like "*ch*2*") Then
            m_match_flag = True
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, X)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, X)
            
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 1)
        End If

    ElseIf (m_siteCnt = 4) Then
        If (m_chmapName Like "*ch*4*") Then
            m_match_flag = True
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, X)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, X)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, X)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(3, X)
            
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y - 4)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 6)
        End If

    ElseIf (m_siteCnt = 6) Then
        If (m_chmapName Like "*ch*6*") Then
            m_match_flag = True
             If UCase(m_chmapName) Like "*CP*" Then
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, X)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, X)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, X + 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(3, X + 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(4, X + 4)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(5, X + 4)
                
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 4)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 4)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(4, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 4)
                
                Debug.Print "Site0: " & "(" & X & "," & y & ")"
                Debug.Print "Site1: " & "(" & X & "," & y - 4 & ")"
                Debug.Print "Site2: " & "(" & X + 2&; "," & y & ")"
                Debug.Print "Site3: " & "(" & X + 2&; "," & y - 4 & ")"
                Debug.Print "Site4: " & "(" & X + 4&; "," & y & ")"
                Debug.Print "Site5: " & "(" & X + 4&; "," & y - 4 & ")"
            ElseIf (UCase(m_chmapName) Like "*WLFT*") Then
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, X)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, X)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, X + 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(3, X + 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(4, X + 4)
                Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(5, X + 4)
                
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 2)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(4, y)
                Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 2)
                
                Debug.Print "Site0: " & "(" & X & "," & y & ")"
                Debug.Print "Site1: " & "(" & X & "," & y - 2 & ")"
                Debug.Print "Site2: " & "(" & X + 2&; "," & y & ")"
                Debug.Print "Site3: " & "(" & X + 2&; "," & y - 2&; ")"
                Debug.Print "Site4: " & "(" & X + 4&; "," & y & ")"
                Debug.Print "Site5: " & "(" & X + 4&; "," & y - 2 & ")"
            
            
            End If
            
        End If
    
    ElseIf (m_siteCnt = 8) Then
        If (m_chmapName Like "*ch*8*") Then
            m_match_flag = True
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, X)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, X)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, X)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(3, X)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(4, X + 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(5, X + 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(6, X + 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(7, X + 2)
            
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y - 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y - 4)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 6)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(4, y)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 2)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(6, y - 4)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(7, y - 6)
        End If

    Else
        m_match_flag = False
    End If
    
    If (m_match_flag = False) Then
        ''''Has the reminder for user to maintain this fuction if the setup is unsuitable.
        TheExec.AddOutput "<WARNING> " + funcName + ":: The Condition Setup is Wrong."
        GoTo errHandler
    End If
    
Exit Function

errHandler:
    TheExec.AddOutput "<Error> " + funcName + ":: please check it out."
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function ReadWaferData()

    Dim LotTmp As String, WaferTmp As String
    Dim X_Tmp As String, Y_Tmp As String
    Dim Loc_dash As Integer
    Dim site As Variant
    
    On Error GoTo err1
    
    '=== Initialization of parameters ====
    
    '=== Simulated Data ===
    If (TheExec.TesterMode = testModeOffline) Then
        LotTmp = "N99G19-01E0"
    Else
        LotTmp = TheExec.Datalog.Setup.LotSetup.LotID

    End If
    
    Loc_dash = InStr(1, LotTmp, "-")
    
    If Loc_dash <> 0 Then
        LotID = Mid(LotTmp, 1, Loc_dash - 1)
    Else
        LotID = LotTmp
    End If
    
    If (TheExec.TesterMode = testModeOffline) Then
        WaferID = Mid(LotTmp, Loc_dash + 1, 2)
    Else
        If TheExec.Datalog.Setup.WaferSetup.ID <> "" Then WaferID = TheExec.Datalog.Setup.WaferSetup.ID
    End If
    
    For Each site In TheExec.sites
        If (TheExec.TesterMode = testModeOffline) Then
            XCoord(site) = 1
            YCoord(site) = 11 - site
        Else
            XCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetXCoord(site)
            YCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetYCoord(site)
                       
            If (XCoord(site) = -32768) Then
                XCoord(site) = 1
                YCoord(site) = 11 - site
            End If
        End If
        
        TheExec.Datalog.WriteComment "Lot ID = " + LotID
        TheExec.Datalog.WriteComment "Wafer ID = " + CStr(WaferID)
        TheExec.Datalog.WriteComment "X coor (site " + CStr(site) + ")= " + CStr(XCoord(site))
        TheExec.Datalog.WriteComment "Y coor (site " + CStr(site) + ")= " + CStr(YCoord(site))
    Next site
    Exit Function
err1:
    TheExec.Datalog.WriteComment ("There is an error happened in the function of ReadWaferData()")
                If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RegKeySave(i_RegKey As String, i_Value As String, Optional i_Type As String = "REG_SZ")
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
End Function
Function RegKeyRead(i_RegKey As String) As String

Dim myWS As Object

On Error Resume Next

Set myWS = CreateObject("WScript.Shell")

RegKeyRead = myWS.RegRead("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\" & i_RegKey)

End Function
Public Function Dec2Bin(ByVal n As Long, ByRef BinArray() As Long)

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

End Function


Public Function Dec2BinStr32Bit(ByVal Nbit As Long, ByVal num As Long) As String
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
End Function
Public Function BinStr2HexStr(ByVal BinStr As String, ByVal HexBit As Long) As String

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

End Function
Function Bin2Dec(sMyBin As String) As Long
    Dim X As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For X = 0 To iLen
        Bin2Dec = Bin2Dec + Mid(sMyBin, iLen - X + 1, 1) * 2 ^ X
    Next
End Function

Function Bin2Dec_rev(sMyBin As String) As Variant
    Dim X As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For X = 0 To iLen
        Bin2Dec_rev = Bin2Dec_rev + Mid(sMyBin, iLen - X + 1, 1) * 2 ^ (iLen - X)
    Next
End Function

Function Bin2Dec_rev_Double(sMyBin As String) As Double
    Dim X As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For X = 0 To iLen
        Bin2Dec_rev_Double = Bin2Dec_rev_Double + Mid(sMyBin, iLen - X + 1, 1) * 2 ^ (iLen - X)
    Next
End Function


Public Function ExculdePath(Pat As Variant) As String
Dim patt_ary_temp() As String
    patt_ary_temp = Split(Pat, "\")
    ExculdePath = patt_ary_temp(UBound(patt_ary_temp))

End Function

''debug printing

Public Function DebugPrintFunc(Test_Pattern As String, Optional testname_enable As Boolean = False) As Long
'for debug printing generation
    Dim PinCnt As Long, PinAry() As String
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
    Dim Irange As Double
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
    
    On Error GoTo errHandler
    
    'version history
    'DebugPrint_version = 1.3   'copy from Fiji
    'DebugPrint_version = 1.4   'implement offline simulation for Rhea bring up
    'DebugPrint_version = 1.5   'Update for Multi-Port nWire setting
     DebugPrint_version = 1.6   'Add differential nWire frequency capture, DCVS tl* modes put in strings, support no pattern items
     DebugPrint_version = 1.7   'Add DC/AC cetegory setup, remove off-limt timing simulation, offline could get real timing.
     DebugPrint_version = 1.71   'Add PPMU debug print function.
     DebugPrint_version = 1.72   'Add DCVI debug print support.
     Shmoo_Pattern = Test_Pattern
     m1_InstanceName = LCase(TheExec.DataManager.instanceName)
    
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
            TheExec.Datalog.WriteComment "  TestInstanceName = " & TheExec.DataManager.instanceName
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

        TheExec.DataManager.DecomposePinList AllPowerPinlist, PinAry(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then
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
                        SlotType = LCase(GetInstrument(PinAry(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Value
                            Case "vhdvs": PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Value
                            Case "dc-07": PowerVolt = TheHdw.DCVI.Pins(PinAry(i)).Voltage
                        End Select
'                    End If
                        
                    TheExec.Datalog.WriteComment "  " & PinAry(i) & " = " & Format(PowerVolt, "0.000") & " v"
                End If
            Next i

            TheExec.Datalog.WriteComment "***** List all power end ******"
            
            TheExec.Datalog.WriteComment "***** List all Vmain power Start ******"

            TheExec.DataManager.DecomposePinList AllPowerPinlist, PinAry(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then
                        SlotType = LCase(GetInstrument(PinAry(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Main.Value
                            Case "vhdvs": PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Main.Value
                            Case "dc-07": PowerVolt = TheHdw.DCVI.Pins(PinAry(i)).Voltage
                        End Select
                    TheExec.Datalog.WriteComment "  " & PinAry(i) & " = " & Format(PowerVolt, "0.000") & " v"
                End If
            Next i

            TheExec.Datalog.WriteComment "***** List all Vmain power end ******"
        
            TheExec.Datalog.WriteComment "***** List all Valt power Start ******"

            TheExec.DataManager.DecomposePinList AllPowerPinlist, PinAry(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then
                        SlotType = LCase(GetInstrument(PinAry(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Alt.Value
                            Case "vhdvs": PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Alt.Value
                            Case "dc-07": PowerVolt = TheHdw.DCVI.Pins(PinAry(i)).Voltage
                        End Select
                    TheExec.Datalog.WriteComment "  " & PinAry(i) & " = " & Format(PowerVolt, "0.000") & " v"
                End If
            Next i

            TheExec.Datalog.WriteComment "***** List all Valt power end ******"
            
            TheExec.Datalog.WriteComment "***** List all power FoldLimit TimeOut Start ******"

            TempString = "FoldLimit TimeOut :"
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'PowerAlramTime = 0.001 * i
'                        PowerAlramTime = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.TimeOut
'                    Else    'online
                        SlotType = LCase(GetInstrument(PinAry(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerAlramTime = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.TimeOut
                            Case "vhdvs": PowerAlramTime = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.TimeOut
                            Case "dc-07": PowerAlramTime = TheHdw.DCVI.Pins(PinAry(i)).FoldCurrentLimit.TimeOut
                        End Select
'                    End If

                    If i <> (PinCnt - 1) Then
                        TempString = TempString + "  " & PinAry(i) & " = " & Format(1000 * PowerAlramTime, "0.000") & " ms" + ","
                    Else
                        TempString = TempString + "  " & PinAry(i) & " = " & Format(1000 * PowerAlramTime, "0.000") & " ms"
                    End If
                End If
            Next i

            TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power FoldLimit TimeOut End ******"
            TheExec.Datalog.WriteComment "***** List all power FoldLimit Current Start ******"

            TempString = "FoldLimit Current :"
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        TempStringOffline = PinAry(i) & "_Ifold_GLB"
'                        Irange = TheExec.specs.Globals(TempStringOffline).ContextValue
'                        'Powerfoldlimit = Irange
'                        Powerfoldlimit = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.Level.Value
'                    Else    'online
                        SlotType = LCase(GetInstrument(PinAry(i), 0))
                        Select Case SlotType
                            Case "hexvs": Powerfoldlimit = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.Level.Value
                            Case "vhdvs": Powerfoldlimit = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.Level.Value
                            Case "dc-07": Powerfoldlimit = TheHdw.DCVI.Pins(PinAry(i)).current
                        End Select
'                    End If

                    If i <> (PinCnt - 1) Then
                        TempString = TempString + "  " & PinAry(i) & " = " & Format(Powerfoldlimit, "0.000000") & " A" + ","
                    Else
                        TempString = TempString + "  " & PinAry(i) & " = " & Format(Powerfoldlimit, "0.000000") & " A"
                    End If
                End If
            Next i

            TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power FoldLimit Current End ******"
            TheExec.Datalog.WriteComment "***** List all power Alram Check Start ******"

            TempString = "Alram Check :"
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'AlarmBehavior = tlAlarmDefault
'                        AlarmBehavior = TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout)
'                    Else
                        SlotType = LCase(GetInstrument(PinAry(i), 0))
                        Select Case SlotType
                            Case "hexvs": AlarmBehavior = TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout)
                            Case "vhdvs": AlarmBehavior = TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout)
                            Case "dc-07": AlarmBehavior = TheHdw.DCVI.Pins(PinAry(i)).FoldCurrentLimit.Behavior
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
                        TempString = TempString + "  " & PinAry(i) & " = " & AlramCheck & ","
                    Else
                        TempString = TempString + "  " & PinAry(i) & " = " & AlramCheck
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
                If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'PowerConnect_State = tlDCVSConnectForce
'                        PowerConnect_State = TheHdw.DCVS.Pins(PinAry(i)).Connected
'                    Else
                        SlotType = LCase(GetInstrument(PinAry(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerConnect_State = TheHdw.DCVS.Pins(PinAry(i)).Connected
                            Case "vhdvs": PowerConnect_State = TheHdw.DCVS.Pins(PinAry(i)).Connected
                            Case "dc-07": PowerConnect_State = TheHdw.DCVI.Pins(PinAry(i)).Connected
                        End Select
'                    End If
                    
                    Select Case PowerConnect_State
                         Case tlDCVSConnectDefault: PowerConnect_State_str = "tlDCVSConnectDefault"
                         Case tlDCVSConnectNone: PowerConnect_State_str = "tlDCVSConnectNone"
                         Case tlDCVSConnectForce: PowerConnect_State_str = "tlDCVSConnectForce"
                         Case tlDCVSConnectSense: PowerConnect_State_str = "tlDCVSConnectSense"
                    End Select
                    If i <> (PinCnt - 1) Then
                        TempString = TempString + "  " & PinAry(i) & " = " & PowerConnect_State_str + ","
                    Else
                        TempString = TempString + "  " & PinAry(i) & " = " & PowerConnect_State_str
                    End If
                End If
            Next i

            TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power Connection Check End ******"
            TheExec.Datalog.WriteComment "***** List all power Gate Start ******"

            TempString = "Power Gate Status:"

            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'Gate_State = True
'                        Gate_State = TheHdw.DCVS.Pins(PinAry(i)).Gate
'                    Else
                       SlotType = LCase(GetInstrument(PinAry(i), 0))
                        Select Case SlotType
                            Case "hexvs": Gate_State = TheHdw.DCVS.Pins(PinAry(i)).Gate
                            Case "vhdvs": Gate_State = TheHdw.DCVS.Pins(PinAry(i)).Gate
                            Case "dc-07": Gate_State = TheHdw.DCVI.Pins(PinAry(i)).Gate
                        End Select
'                    End If
                    
                    Select Case Gate_State
                         Case True: Gate_State_str = "on"
                         Case False: Gate_State_str = "off"
                    End Select
                    If i <> (PinCnt - 1) Then
                        TempString = TempString + "  " & PinAry(i) & " = " & Gate_State_str + ","
                    Else
                        TempString = TempString + "  " & PinAry(i) & " = " & Gate_State_str
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

'                TheExec.Datalog.WriteComment "  Pins : " & CStr(EachPinGroup) _
'                & " , Vih = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVih), "0.000") & " v" _
'                & " , Vil = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVil), "0.000") & " v" _
'                & " , Voh = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVoh), "0.000") & " v" _
'                & " , Vol = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVol), "0.000") & " v" _
'                & " , Iol = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chIoh), "0.000") & " v" _
'                & " , Ioh = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chIol), "0.000") & " v" _
'                & " , Vt  = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVt), "0.000") & " v" _
'                & " , Vch = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVch), "0.000") & " v" _
'                & " , Vcl = " & Format(thehdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVcl), "0.000") & " v" _
'                & " , PPMU_VclampHi = " & Format(thehdw.PPMU.Pins(CStr(EachPinGroup)).ClampVHi, "0.000") & " v" _
'                & " , PPMU_VclampLow = " & Format(thehdw.PPMU.Pins(CStr(EachPinGroup)).ClampVLo, "0.000") & " v"
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
            Dim site As Variant
          
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
                For Each site In TheExec.sites
                    XI0_Freq_pl.Pins(0).Value = 24000000
                Next site
            End If

            For Each site In TheExec.sites
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
            Next site

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
            TheExec.DataManager.DecomposePinList All_Utility_list, PinAry(), PinCnt

            'Utility bits
            out_line = "Utility_list : "
            For Each CurrSite In TheExec.sites.Active
                For i = 0 To PinCnt - 1
                    If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then
                        PinData = TheHdw.Utility.Pins(PinAry(i)).States(tlUBStateProgrammed)    'TheHdw.Utility.pins((pinary(i)) '.States(tlUBStateCompared)
                        If i = 0 Then
                              out_line = out_line + PinAry(i) & " = " & PinData.Pins(0).Value(CurrSite) '''& ","
                        Else
                              out_line = out_line & "," & PinAry(i) & " = " & PinData.Pins(0).Value(CurrSite)
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
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function FreeRunClk_ScopeOut(PAPort As PinList, Optional DebugFlag As Boolean = False)

    On Error GoTo errHandler
    Dim TempStr As String

    TheHdw.Digital.Pins(PAPort).Disconnect

    If DebugFlag = True Then
        TheExec.Datalog.WriteComment "print: nWire scope out, port (" & PAPort.Value & ")"
    End If

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function
Public Function FreeRunClk_ScopeIn(PAPort As PinList, Optional DebugFlag As Boolean = False) ''update for multi nWire 20170718

    On Error GoTo errHandler
    Dim TempStr As String
    
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
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next


End Function

Public Function PowerUp_Interpose(PAPort As PinList, Optional DebugFlag As Boolean = False)
    On Error GoTo errHandler
    
'    FreeRunClk_ScopeOut PAPort, DebugFlag
    'TheHdw.Utility.Pins(Relay).State = tlUtilBitOn
    
'    If DebugFlag = True Then    'debugprint
'         TheExec.Datalog.WriteComment "print: RTCLK relay on, relay " & Relay.Value
'    End If

    FreeRunClk_ScopeIn PAPort, DebugFlag
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function PowerDown_Interpose(nWireDisconnectPin As String, Optional DebugFlag As Boolean = False)
    On Error GoTo errHandler

    FreeRunClk_Disconnect nWireDisconnectPin, DebugFlag
'    FreeRunClk_Disable nWireDisconnectPin, True 'pass site will halt also
    TheExec.Datalog.WriteComment "print: nWire engine , Halt " & vbCrLf

    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Function IEDA_Initialize(ByRef InputStr As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "IEDA_Initialize"
    
    InputStr = ""

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Function IEDA_SaveRegistry(ByVal InputStr As String, RegistryName As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "IEDA_SaveRegistry"

    Call RegKeySave(RegistryName, InputStr)
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Bintable_initial()
'//=====================================================================================
    
On Error GoTo errHandler

    Dim BintableSheet As Worksheet
    Dim BinNameColumnMax As Long
    Dim BinColumnNum As Long
    Dim BinContext As String
    Dim BinColNumAccu As Long: BinColNumAccu = 0
    
    '20170529 add sheet loop to include all bin table sheets
    For Each BintableSheet In ThisWorkbook.Sheets
        If LCase(BintableSheet.Name) Like "*bin*table*" Then
            #If IGXL8p30 Then
            #Else
                BintableSheet.Activate
            #End If
            BinNameColumnMax = BintableSheet.Cells(Rows.Count, 2).End(xlUp).row
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
            
    For lCount = 0 To UBound(tyBinTable.astrBinName)
        If tyBinTable.astrBinSortNum(lCount) <> "" Then
            If tyBinTable.astrBinRename(lCount) <> "" Then
                Call TheExec.Datalog.SBRFill(tyBinTable.astrBinSortNum(lCount), tyBinTable.astrBinRename(lCount))
            Else
                Call TheExec.Datalog.SBRFill(tyBinTable.astrBinSortNum(lCount), Mid(tyBinTable.astrBinName(lCount), InStr(tyBinTable.astrBinName(lCount), "_") + 1))
            End If
        End If
    Next lCount

    Exit Function

errHandler:


'//============================================================================================
                If AbortTest Then Exit Function Else Resume Next
End Function
Public Function DebugPrintFunc_PPMU(PPMU_Pins As String) As Long
'for debug printing generation
    Dim PinCnt As Long, PinAry() As String
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
    Dim PPMU_Forcei As String
    Dim DCVI_Mode As String
    Dim DCVI_sense_relay As Boolean
    Dim DCVI_force_relay As Boolean
    
    On Error GoTo errHandler
    
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
        TheExec.Datalog.WriteComment "  TestInstanceName = " & TheExec.DataManager.instanceName
        TheExec.Datalog.WriteComment "***** List all power Start ******"

        TheExec.DataManager.DecomposePinList AllPowerPinlist, PinAry(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then

                        PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Main.Value
                        
                    TheExec.Datalog.WriteComment "  " & PinAry(i) & " = " & Format(PowerVolt, "0.000") & " v"
                End If
            Next i
        TheExec.Datalog.WriteComment "***** List all power end ******"



        TheExec.Datalog.WriteComment "***** List all DCVI Start ******"

        TheExec.DataManager.DecomposePinList AllDCVIPinlist, PinAry(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(PinAry(i)) <> "N/C" Then

                    PowerVolt = TheHdw.DCVI.Pins(PinAry(i)).Voltage
                    PowerCurrent = TheHdw.DCVI.Pins(PinAry(i)).current
                    
                    If TheHdw.DCVI.Pins(PinAry(i)).mode = tlDCVIModeVoltage Then
                        DCVI_Mode = "ForceV"
                    ElseIf TheHdw.DCVI.Pins(PinAry(i)).mode = tlDCVIModeCurrent Then
                        DCVI_Mode = "ForceI"
                    Else
                        DCVI_Mode = "HighImpedance"
                    End If


                    If TheHdw.DCVI.Pins(PinAry(i)).Connected = 0 Then
                        DCVI_force_relay = False
                        DCVI_sense_relay = False
                    ElseIf TheHdw.DCVI.Pins(PinAry(i)).Connected = 1 Then
                        DCVI_force_relay = True
                        DCVI_sense_relay = False
                    ElseIf TheHdw.DCVI.Pins(PinAry(i)).Connected = 2 Then
                        DCVI_force_relay = False
                        DCVI_sense_relay = True
                    ElseIf TheHdw.DCVI.Pins(PinAry(i)).Connected = 3 Then
                        DCVI_force_relay = True
                        DCVI_sense_relay = True
                    End If

                    TheExec.Datalog.WriteComment "  DCVI_Pins : " & PinAry(i) _
                    & " , Voltage = " & Format(PowerVolt, "0.000000") & " v" _
                    & " , Current = " & Format(PowerCurrent, "0.000000") & " A" _
                    & " , Mode = " & DCVI_Mode & " " _
                    & " , Gate = " & TheHdw.DCVI.Pins(PinAry(i)).Gate _
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
                        PPMU_Forcei = CStr(Format(TheHdw.PPMU.Pins(PPMU_used_Pin).current.Value, "0.000000"))
                        
                        If TheHdw.PPMU.Pins(CStr(PPMU_used_Pin)).mode = tlPPMUForceVMeasureI Then
                            PPMU_Forcei = "None"
                        Else
                            PPMU_ForceV = "None"
                        End If
        
                        TheExec.Datalog.WriteComment "  Pins : " & CStr(PPMU_used_Pin) _
                        & " , PPMU_VclampHi = " & Format(TheHdw.PPMU.Pins(CStr(PPMU_used_Pin)).ClampVHi, "0.000") & " v" _
                        & " , PPMU_VclampLow = " & Format(TheHdw.PPMU.Pins(CStr(PPMU_used_Pin)).ClampVLo, "0.000") & " v" _
                        & " , PPMU_forceV = " & PPMU_ForceV & " v" _
                        & " , PPMU_ForceI = " & PPMU_Forcei & " A"
                  End If
                Next PPMU_used_Pin
            End If
        
        TheExec.Datalog.WriteComment "***** List PPMU condition end ******"
            

            TheExec.Datalog.WriteComment "================debug print PPMU end  =================="
            TheExec.Datalog.WriteComment ""
        End If
    Exit Function
    
errHandler:
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
'    On Error GoTo errhandler
    
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
errHandler:
    GetPatFromPatternSet = False
    rtnPatCnt = -1

                If AbortTest Then Exit Function Else Resume Next
End Function
Public Function FreeRunClk_Disconnect(nWireDisconnectPin As String, Optional DebugFlag As Boolean = False)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "FreeRunClk_Disconnect"

    TheHdw.Digital.Pins(nWireDisconnectPin).Disconnect
    
    If DebugFlag = True Then TheExec.Datalog.WriteComment "print: nWire disconnect, pin " & nWireDisconnectPin

    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Find_nWire_Pin() As Long  ''Support multiple nWire port 20170718
' Get all nWire port and put in global variable nWire_Ports_GLB
    Dim i As Long
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim row_cnt As Long
    Dim nWire_cnt As Long
    Dim nWire_Pin_ary(10) As String
    Dim curr_pin As String, last_pin As String
    
'    If nWire_Ports_GLB <> "" Then Exit Function
    nWire_Ports_GLB = ""
    
    Set wb = Application.ActiveWorkbook
    Set ws = wb.Sheets("Levels_nWire")
    ws.Activate
    
'    row_cnt = ws.Cells(Rows.Count, 2).End(xlUp).Row - 1
    nWire_cnt = 0
    For i = 4 To Rows.Count 'skip header line
        curr_pin = ws.Cells(i, 2)
        If curr_pin = "" Then
            i = Rows.Count + 1 'stop at empty row/cell
        ElseIf curr_pin Like "*_PA" And curr_pin <> last_pin Then
            nWire_Pin_ary(nWire_cnt) = curr_pin
            last_pin = curr_pin
            nWire_cnt = nWire_cnt + 1
        End If
    Next i
    For i = 1 To nWire_cnt
        If nWire_Ports_GLB <> "" Then
            nWire_Ports_GLB = nWire_Ports_GLB & "," & nWire_Pin_ary(i - 1)
        Else
            nWire_Ports_GLB = nWire_Pin_ary(i - 1)
        End If
    Next i
End Function

Public Function Get_nWire_Name(NWire As Variant, port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_PowerSequence_pa As String) ''Support multiple nWire port 20170718
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
End Function

Public Function Disable_FRC(nWire_ports As String, Optional DisConnectFRC As Boolean = False) ''Support multiple nWire port 20170718
' nWire_ports  can be port name or pin name
' If it is blank, will assume to use all nWire ports
    'Eg. nWire_ports = "XI0_Port, RT_CLK32768_Port, XIN_Port"
    Dim nWire_port_ary() As String
    Dim nwp As Variant, all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim site As Variant
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
    For Each site In TheExec.sites
        TheHdw.Protocol.ports(all_ports).Halt
        TheHdw.Protocol.ports(all_ports).Enabled = False
    Next site
    
    If DisConnectFRC = True Then
        TheExec.Datalog.WriteComment "******************  Disconnect nWire pins " & all_pins & " ****************"
        TheHdw.Digital.Pins(all_pins).Disconnect
    End If
End Function

Public Function Enable_FRC(nWires As String, Optional ConnectFRC As Boolean = False) ''Support multiple nWire port 20170718
' nWires  can be port name or pin name
' If it is blank, will assume to use all nWire ports
    'Eg. nWires = "XI0_Port, RT_CLK32768_Port, XIN_Port"
    Dim nWires_ary() As String
    Dim nwp As Variant, all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim PLL_Lock As New SiteLong
    Dim port_level_value As Double
    Dim FreeRunFreq As Double
''''Start, modify from T-sic, Carter, 20191022
    Dim Flag_IsPLLLocked As Boolean
''''End, modify from T-sic, Carter, 20191022
    Dim site As Variant

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
''''Start, modify from T-sic, Carter, 20191022
    nWires_ary = Split(all_ports, ",")
    For Each nwp In nWires_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        If TheHdw.Protocol.ports(nwp).Family = "FRC" Then
            TheHdw.Protocol.ports(nwp).FRC.ResetPLL
            TheHdw.Wait 0.001
            If TheHdw.Protocol.ports(nwp).Status = tlProtocolPortStatus_Running Then
            Else
                Call TheHdw.Protocol.ports(nwp).FRC.start
            End If
        Else '' nWire
            TheHdw.Protocol.ports(nwp).NWire.ResetPLL
            TheHdw.Wait 0.001
            Call TheHdw.Protocol.ports(nwp).NWire.Frames("RunFreeClock").Execute
            TheHdw.Protocol.ports(nwp).IdleWait
        End If
    Next nwp
    
''''End, modify from T-sic, Carter, 20191022

    TheExec.Datalog.WriteComment "Enable nWire Clock " & all_ports
    '****print out to data log about nWire clock condition
    nWires_ary = Split(all_ports, ",")
    For Each nwp In nWires_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        
        For Each site In TheExec.sites
''''Start, modify from T-sic, Carter, 20191022
            If TheHdw.Protocol.ports(nwp).Family = "FRC" Then
                Flag_IsPLLLocked = TheHdw.Protocol.ports(nwp).FRC.IsPLLLocked
            Else '' nWire
                Flag_IsPLLLocked = TheHdw.Protocol.ports(nwp).NWire.IsPLLLocked
            End If
            
            If Flag_IsPLLLocked = False Then
                PLL_Lock = 0
            Else
                PLL_Lock = 1
            End If
''''End, modify from T-sic, Carter, 20191022
        Next site
        
        FreeRunFreq = 1 / TheHdw.Digital.Timing.Period(nwp) / 1000000
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites.Selected
                PLL_Lock = 1
                FreeRunFreq = TheExec.specs.AC.Item(ac_spec_pa).CurrentValue / 1000000  'offline
             Next site
        End If
    
        
        TheExec.Flow.TestLimit PLL_Lock, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="nWire " & nwp & " PLL_Lock" 'BurstResult=1:Pass
        If LCase(nwp) Like "*diff*" Then
            port_level_value = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.Value(chVid)
            TheExec.Datalog.WriteComment "********** freerunning clock(" & nwp & ") = " & Format(FreeRunFreq, "0.000") & " Mhz, Vid = " & port_level_value
        Else
            port_level_value = TheHdw.Digital.Pins(pin_pa).Levels.Value(chVih)
            TheExec.Datalog.WriteComment "********** freerunning clock(" & nwp & ") = " & Format(FreeRunFreq, "0.000") & " Mhz, Vih = " & port_level_value
        End If
    Next nwp
End Function

Public Function Meas_FRC(nWire_ports As String)
    'Eg. nWire_ports = "XI0_Port, RT_CLK32768_Port, XIN_Port"
    Dim nWire_port_ary() As String
    Dim nwp As Variant, meas_freq As New PinListData, site As Variant
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
            For Each site In TheExec.sites
                meas_freq.Pins(0).Value = TheExec.specs.AC(ac_spec_pa).CurrentValue
            Next site
        End If
                        
        For Each site In TheExec.sites
            If port_pa Like "*DIFF*" Then
                PA_Vicm = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.Value(chVicm)
                PA_Vid = TheHdw.Digital.Pins(pin_pa).DifferentialLevels.Value(chVid)
                PA_Vihd = PA_Vicm + PA_Vid / 2
                PA_Vild = PA_Vicm - PA_Vid / 2
                TheExec.Datalog.WriteComment "  FreeRunFreq (" & pin_pa & ") : " & Format(meas_freq.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(PA_Vihd, "0.000") & " v , clock_Vil: " & Format(PA_Vild, "0.000") & " v"
            Else
                TheExec.Datalog.WriteComment "  FreeRunFreq (" & pin_pa & ") : " & Format(meas_freq.Pins(pin_pa).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(pin_pa).Levels.Value(chVih), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(pin_pa).Levels.Value(chVil), "0.000") & " v"
            End If

        Next site
    Next nwp
End Function

Public Function StartProfile(PinName As String, WhatToCapture As String, SampleRate As Double, SampleSize As Double, CapSignalName As String, Slot_Type As String, _
Optional Meter_I_Range As Double, Optional Meter_V_Range As Double)

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
            .mode = tlDCVSMeterCurrent
            '.Range = Meter_I_Range
        Else
            .mode = tlDCVSMeterVoltage
            '.Range = Meter_V_Range
        End If
        .SampleRate = SampleRate
        .SampleSize = SampleSize

    End With

    ' Setup the hardware by loading the signal
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName).LoadSettings

    ' Start the capture
    TheHdw.DCVS.Pins(PinName).Capture.Signals.Item(CapSignalName).Trigger

End Function

Public Function SplitPinByinstrument(PinName As String, ByRef HexPins As String, ByRef UVSPins As String)

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

End Function

Public Function ProfileAutoResolution(SlotType As String, measuretime As Double, ByRef SampleSize As Double, ByRef SampleRate As Double)

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
    
End Function

Function Bin2Dec_rev_Fractional(sMyBin As String) As Variant
    Dim X As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For X = 0 To iLen
        Bin2Dec_rev_Fractional = Bin2Dec_rev_Fractional + Mid(sMyBin, iLen - X + 1, 1) * 2 ^ (-X - 1)
    Next
End Function

Public Function Dec2Bin_str(ByVal n As Long, ByRef BinArray_string As String, bitCount As Long)

    Dim i As Integer, j As Integer
    Dim Element_Amount As Integer
    Dim Count As Integer
    Dim BinArray() As Long
    BinArray_string = ""
    ReDim BinArray(bitCount) As Long
    
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
    For i = 0 To UBound(BinArray)
        BinArray_string = BinArray_string & BinArray(UBound(BinArray) - i)
    Next i

End Function
Public Function ShmooEndFunction() As Boolean

    If TheExec.DevChar.Setups.IsRunning = True Then
        Dim site As Variant
        Dim SetupName As String
        Dim X_RangeFrom As Double
        Dim Y_RangeFrom As Double
        
        SetupName = TheExec.DevChar.Setups.ActiveSetupName
        If Not ((TheExec.DevChar.Results(SetupName).startTime Like "1/1/0001*" Or TheExec.DevChar.Results(SetupName).startTime Like "0001/1/1*")) Then
            gl_Flag_HardIP_Characterization_1stRun = False
            With TheExec.DevChar.Setups(SetupName)
                If .Shmoo.Axes.Count > 1 Then
                    X_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
                    Y_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.from
                    For Each site In TheExec.sites ''20181101 current point need site value
                        XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
                        YVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
                    Next site
                    If XVal = X_RangeFrom And YVal = Y_RangeFrom Then
                        gl_flag_end_shmoo = False
                    End If
                    If gl_flag_end_shmoo = True Then
                        ShmooEndFunction = True
                    End If
                Else
                    X_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
                    For Each site In TheExec.sites ''20181101 current point need site value
                        XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
                    Next site
'                    If XVal = X_RangeFrom Then
'                        gl_flag_end_shmoo = False
'                    End If
                    If gl_flag_end_shmoo = True Then
                        ShmooEndFunction = True
                    End If
                End If
            End With
        Else
           gl_Flag_HardIP_Characterization_1stRun = True
           gl_flag_end_shmoo = False
        End If
    End If

End Function

