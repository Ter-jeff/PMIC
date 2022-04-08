Attribute VB_Name = "LIB_Vdd_Binning_customer"
Option Explicit

'20200320: Modified to check instance contexts of current instance and previous instance.
'20191127: Modified for the revised InitVddBinTable.
Public Function init_before_patt_customize(inst_name As String, p_mode As Integer, passBinCut As SiteLong, Optional Local_KeepAliveFlag As Boolean = False)
    Dim strTemp As String
On Error GoTo errHandler
    strTemp = LCase(inst_name)

    If strTemp Like "*spi*" Or strTemp Like "*rtos*" Then
        If Local_KeepAliveFlag = True And strTemp Like "*ka*" Then
            TheHdw.Digital.Patgen.KeepAlive.Enable = False
            TheHdw.Digital.Patgen.Halt
        End If
        
        If strTemp Like "*ka*" Then
            '''//Check instance contexts of current instance and previous instance.
            CurrentBinCutInstanceContext = GetInstanceContextIntoString
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            IsLevelLoadedForApplyLevelsTiming = True
            PreviousBinCutInstanceContext = CurrentBinCutInstanceContext
        End If
        
        '''20191118: Leon removed the vbt code.
        'Call SPI_Initial_Conds_Fun
        
        If Local_KeepAliveFlag = True And strTemp Like "*ka*" Then
            TheHdw.Digital.Patgen.KeepAlive.Flag = cpuB
            TheHdw.Digital.Patgen.KeepAlive.Enable = True
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of init_before_patt_customize"
    If AbortTest Then Exit Function Else Resume Next
End Function

'''Public Function init_before_patt_payload(inst_name As String)
'''    Dim StrTemp As String
'''On Error GoTo errHandler
'''    StrTemp = LCase(inst_name)
'''
'''    If StrTemp Like "*spi*" Or StrTemp Like "*rtos*" Then
'''        Call SPI_inital_Before_Pat
'''    End If
'''    TheHdw.Wait 0.001
'''Exit Function
'''errHandler:
'''    TheExec.Datalog.WriteComment "Error encountered in VBT Function of init_before_patt_payload"
'''    If AbortTest Then Exit Function Else Resume Next
'''End Function

'20190507: Chris and Verity masked the code for reducing 9.1 P5 compiler error because they thought Oasis might cause the compiler crashed.
Public Function UserSiteVar(responseID As String, upperbound As Variant, lowerbound As Variant) As SiteVariant
    Dim mySAV As New SiteVariant
On Error GoTo errHandler
    '''//Make the measurement Read
    Set mySAV = GenerateRandomSiteVar(upperbound, lowerbound)
    Set UserSiteVar = mySAV
Exit Function
errHandler:
    TheExec.ErrorLogMessage "Error: " + err.Description + " Source: " + err.Source
    If AbortTest Then Exit Function Else Resume Next
End Function

'Public Function IGSim_IncreaseResponse_GenerateScenario(ResponseSheetName As String, StartRow As Long, AddTimes As Long, ScenarioSheetName As String) As Long
'    Dim ws_def As Worksheet
'    Dim wb As Workbook
'    Dim Row As Long
'    Dim string_split() As String
'    Dim id_split() As String
'    Dim temp_string As String
'    Dim Array_size As Long
'    Dim response_content() As String
'    Dim scenario_content() As String
'    Dim stored_row As Long
'    Dim new_response_id As String
'    Dim i As Integer, j As Integer
'    Dim gen_scenario_enable As Boolean
'    Dim test_ins_count As Integer
'    Dim last_test_ins_name As String
'    Dim test_ins_string As String
'    Dim test_ins_name() As String
'    Dim non_repeat_response() As String
'    Dim new_scenario_row As Integer
'On Error GoTo errHandler
'    Set wb = Application.ActiveWorkbook
'    Set ws_def = wb.Sheets(ResponseSheetName)
'
'    If ScenarioSheetName <> "" Then
'        gen_scenario_enable = True
'        new_scenario_row = 0
'        test_ins_count = 0
'        last_test_ins_name = ""
'        test_ins_string = ""
'        ReDim non_repeat_response(StartRow - 4)
'
'        For i = 3 To StartRow - 1
'            non_repeat_response(i - 3) = ws_def.Cells(i, 1).Value
'        Next i
'    End If
'
'    stored_row = 0
'    Array_size = range("A" & Rows.Count).End(xlUp).Row
'    Row = StartRow
'    Array_size = (Array_size - StartRow + 1) * AddTimes
'
'    ReDim response_content(Array_size, 6)
'
'    While ws_def.Cells(Row, 1).Value <> ""
'        temp_string = ws_def.Cells(Row, 1).Value
'        string_split() = Split(temp_string, ".")
'
'        If UBound(string_split) > 2 Then
'            '''Do nothing
'        Else
'            If string_split(1) = 0 Then
'                id_split() = Split(string_split(0), "_")
'                If id_split(UBound(id_split)) = "PT" Then
'                    If id_split(0) = "GSHV" Then
'                        If gen_scenario_enable = True Then
'                            If last_test_ins_name <> ws_def.Cells(Row, 5).Value Then
'                                If test_ins_string <> "" Then
'                                    test_ins_string = test_ins_string & "," & ws_def.Cells(Row, 5).Value
'                                Else
'                                    test_ins_string = ws_def.Cells(Row, 5).Value
'                                End If
'                                last_test_ins_name = ws_def.Cells(Row, 5).Value
'                            End If
'                        End If
'                        For i = 0 To AddTimes
'                            temp_string = string_split(0) & "." & i
'                            response_content(stored_row, 0) = temp_string
'                            For j = 1 To 5
'                                response_content(stored_row, j) = ws_def.Cells(Row, j + 1).Value
'                            Next j
'                            stored_row = stored_row + 1
'                        Next i
'                    Else
'                        If gen_scenario_enable = True Then
'                            If last_test_ins_name <> ws_def.Cells(Row, 5).Value Then
'                                If test_ins_string <> "" Then
'                                    test_ins_string = test_ins_string & "," & ws_def.Cells(Row, 5).Value
'                                Else
'                                    test_ins_string = ws_def.Cells(Row, 5).Value
'                                End If
'                                last_test_ins_name = ws_def.Cells(Row, 5).Value
'                            End If
'                        End If
'                        new_response_id = ""
'                        new_response_id = Replace(string_split(0), "PT", "PTBP")
'                        For i = 0 To AddTimes
'                            temp_string = string_split(0) & "." & i
'                            response_content(stored_row, 0) = temp_string
'                            For j = 1 To 5
'                                response_content(stored_row, j) = ws_def.Cells(Row, j + 1).Value
'                            Next j
'                            stored_row = stored_row + 1
'                            temp_string = new_response_id & "." & i
'                            response_content(stored_row, 0) = temp_string
'                            For j = 1 To 5
'                                response_content(stored_row, j) = ws_def.Cells(Row, j + 1).Value
'                            Next j
'                            stored_row = stored_row + 1
'                        Next i
'                    End If
'                End If
'            End If
'        End If
'        Row = Row + 1
'    Wend
'
'    If gen_scenario_enable = True Then
'        test_ins_name() = Split(test_ins_string, ",")
'        test_ins_count = UBound(test_ins_name) + 1
'        Array_size = ((stored_row - 1) + UBound(non_repeat_response) + 1) * test_ins_count
'        ReDim scenario_content(Array_size - 1, 6)
'    End If
'
'    Row = StartRow
'    For i = 0 To stored_row - 1
'        For j = 0 To 5
'            ws_def.Cells(Row, j + 1).Value = response_content(i, j)
'        Next j
'        Row = Row + 1
'    Next i
'
'    If gen_scenario_enable = True Then
'        For i = 0 To test_ins_count - 1
'            For j = 0 To UBound(non_repeat_response) - 1
'                scenario_content(new_scenario_row, 0) = "SC" & i + 1
'                scenario_content(new_scenario_row, 1) = non_repeat_response(j)
'                scenario_content(new_scenario_row, 2) = "Default"
'                new_scenario_row = new_scenario_row + 1
'            Next j
'            For j = 0 To stored_row - 1
'                scenario_content(new_scenario_row, 0) = "SC" & i + 1
'                scenario_content(new_scenario_row, 1) = response_content(j, 0)
'                If response_content(j, 4) <> test_ins_name(i) Then
'                    scenario_content(new_scenario_row, 2) = "PASS"
'                Else
'                    scenario_content(new_scenario_row, 2) = "FAIL"
'                End If
'                new_scenario_row = new_scenario_row + 1
'            Next j
'        Next i
'        Set ws_def = wb.Sheets(ScenarioSheetName)
'        Row = 1
'        For i = 0 To new_scenario_row - 1
'            For j = 0 To 2
'                ws_def.Cells(Row, j + 1).Value = scenario_content(i, j)
'            Next j
'            Row = Row + 1
'        Next i
'    End If
'Exit Function
'errHandler:
'   TheExec.Datalog.WriteComment "Error encountered in VBT Function of IGSim_IncreaseResponse_GenerateScenario"
'   If AbortTest Then Exit Function Else Resume Next
'End Function

Private Function GenerateRandomSiteVar(upperbound As Variant, lowerbound As Variant) As SiteVariant
    Dim site As Variant
    Dim sitevar As New SiteVariant
On Error GoTo errHandler
    For Each site In TheExec.sites.Selected
       'sitevar(site) = (upperbound - lowerbound) * Rnd + lowerbound
        sitevar(site) = (upperbound - lowerbound) * 0 + lowerbound
        If sitevar(site) > upperbound Then
            sitevar(site) = upperbound
        ElseIf sitevar(site) < lowerbound Then
            sitevar(site) = lowerbound
        End If
    Next site
    Set GenerateRandomSiteVar = sitevar
Exit Function
errHandler:
    TheExec.ErrorLogMessage "Error: " + err.Description + " Source: " + err.Source
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210131: Modified to adjust the sequence of parsing SELSRM_Mapping_Table.
'20200922: Modified to check "PRESERVED" and "RESERVED". C651 Toby requested this on 20200916.
'20200709: Modified to check powerDomain names.
'20200703: Modiifed to use "check_Sheet_Range".
'20200528: Modified to check header of the table.
'20200210: As per discussion with C651 Chris and Toby, they will provide patterns with vbump for all BinCut powerDomains.
'20200203: Modified to parse the column "Comment"
'20200129: Modified to check if the row is "End" or empty.
'20200103: Modified to check if Logic Pin is "PRESERVED".
'20200102: Modified for bit sequence with "PRESERVED" in columns "Logic Pins" and "Sram Pins".
'20191120: Modified to add the optional argument to check logicPin with BinCut corePower pinGroup and sramPin with BinCut otherRail pinGroup.
'20191120: Modified to check if the content matches the current TestJob.
'20191113: Modified to check if the pinGroup is empty or not.
'20191017: Modified to parse the pattern keywords.
'20191007: Modified the parsing method for two SELSRM_Mapping_Table sheets by test jobs for Tonga.
'20190906: Modified the parsing method for the different SELSRAM DSSC bit length.
Public Function Parsing_SELSRM_Mapping_Table(testJob As String, Optional bincutCorepowerPingroup As String = "", Optional bincutOtherrailPingroup As String = "") '''20180522 '''
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim i As Long, j As Long, k As Long
    Dim row_of_title As Integer
    Dim col_stage As Integer
    Dim col_Block As Integer
    Dim col_Pattern As Integer
    Dim col_LogicPin As Integer
    Dim col_SramPin As Integer
    Dim col_Selsram1 As Integer
    Dim col_Selsram0 As Integer
    Dim col_BitSeq As Integer
    Dim col_Alpha As Integer
    Dim col_comment As Integer
    Dim tempTestjob As String
    Dim lastTestjob As String
    Dim currentBlock As String
    Dim tempPattern As String
    Dim tempBitSeq As String
    Dim tempLogicPinSeq As String
    Dim tempSramPinSeq As String
    Dim tempSelsram1Seq As String
    Dim tempSelsram0Seq As String
    Dim tempAlphaSeq As String
    Dim tempCommentSeq As String
    '''variables
    Dim blockCnt As Integer
    Dim bitCnt As Integer
    Dim tempPinalpha As String
    Dim currentPattern As String
    Dim foundPin As Boolean
    Dim foundTestjob As Boolean
    Dim split_content() As String
    Dim testjobMatched As String
    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
On Error GoTo errHandler
    If testJob = "" Then
        TheExec.Datalog.WriteComment "The argument of testJob for Parsing_SELSRM_Mapping_Table should not be empty. Error!!!"
        TheExec.ErrorLogMessage "The argument of testJob for Parsing_SELSRM_Mapping_Table should not be empty. Error!!!"
    End If
    
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "SELSRM_Mapping_Table"
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        tempBitSeq = ""
        tempTestjob = ""
        lastTestjob = "None"
        currentBlock = ""
        currentPattern = ""
        tempPattern = ""
        tempBitSeq = ""
        tempLogicPinSeq = ""
        tempSramPinSeq = ""
        tempSelsram1Seq = ""
        tempSelsram0Seq = ""
        tempAlphaSeq = ""
        tempCommentSeq = ""
        selsramPin = ""
        selsramLogicPin = ""
        selsramSramPin = ""
        tempPinalpha = ""
        blockCnt = -1
        bitCnt = -1
        foundTestjob = False
        testjobMatched = ""
        enableRowParsing = False
        
        '''//Check the header of the table
        '''Get the columns for the diverse coefficient
        For i = 1 To MaxRow
            For j = 1 To maxcol
                If UCase(ws_def.Cells(i, j).Value) = "STAGE" Then
                    col_stage = j
                    row_of_title = i
                End If
                
                If UCase(ws_def.Cells(i, j).Value) = "BLOCK" Then
                    col_Block = j
                    row_of_title = i
                End If
                
                If row_of_title > 0 Then
                    If UCase(ws_def.Cells(row_of_title, j).Value) = "STAGE" Then
                        col_stage = j
                    ElseIf UCase(ws_def.Cells(row_of_title, j).Value) = "BITS" Then
                        col_BitSeq = j
                    ElseIf UCase(ws_def.Cells(row_of_title, j).Value) = "PATTERN" Then
                        col_Pattern = j
                    ElseIf UCase(ws_def.Cells(row_of_title, j).Value) = "LOGIC PINS" Then
                        col_LogicPin = j
                    ElseIf UCase(ws_def.Cells(row_of_title, j).Value) = "SRAM PINS" Then
                        col_SramPin = j
                    ElseIf UCase(ws_def.Cells(row_of_title, j).Value) = "SELSRM1" Then
                        col_Selsram1 = j
                    ElseIf UCase(ws_def.Cells(row_of_title, j).Value) = "SELSRM0" Then
                        col_Selsram0 = j
                    ElseIf UCase(ws_def.Cells(row_of_title, j).Value) = "ALPHA" Then
                        col_Alpha = j
                    ElseIf UCase(ws_def.Cells(row_of_title, j).Value) = "COMMENT" Then
                        col_comment = j
                    Else
                        '''Do nothing
                    End If
                End If
            Next j
            
            If row_of_title > 0 Then
                If col_BitSeq > 0 And col_Pattern > 0 And col_LogicPin > 0 And col_SramPin > 0 And col_Selsram1 > 0 And col_Selsram0 > 0 And col_Alpha > 0 And col_comment > 0 Then
                    enableRowParsing = True
                Else
                    enableRowParsing = False
                    
                    If col_stage = 0 Then
                        TheExec.Datalog.WriteComment "Column STAGE doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column STAGE doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_BitSeq = 0 Then
                        TheExec.Datalog.WriteComment "Column BITS doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column BITS doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_LogicPin = 0 Then
                        TheExec.Datalog.WriteComment "Column LOGIC PINS doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column LOGIC PINS doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_SramPin = 0 Then
                        TheExec.Datalog.WriteComment "Column SRAM PINS doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column SRAM PINS doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_Selsram1 = 0 Then
                        TheExec.Datalog.WriteComment "Column SELSRM1 doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column SELSRM1 doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_Selsram0 = 0 Then
                        TheExec.Datalog.WriteComment "Column SELSRM0 doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column SELSRM0 doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_Alpha = 0 Then
                        TheExec.Datalog.WriteComment "Column ALPHA doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column ALPHA doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_comment = 0 Then
                        TheExec.Datalog.WriteComment "Column COMMENT doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column COMMENT doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                End If
                
                Exit For
            End If
        Next i
        
        '''//Start parsing the cells
        If enableRowParsing = True Then
            For i = row_of_title + 1 To MaxRow
                foundTestjob = True
                
                '''//Check if the row is "End" or empty.
                If LCase(ws_def.Cells(i, 1).Value) = "end" Or ws_def.Cells(i, col_BitSeq).Value = "" Then
                    Exit For
                End If
                
                '''//If stageCol exists, check if the content matches the current TestJob.
                '''20210131: Modified to adjust the sequence of parsing SELSRM_Mapping_Table.
                If col_stage > 0 Then '''with column "Stage"
                    If CInt(ws_def.Cells(i, col_BitSeq).Value) = 0 Then
                        tempTestjob = ws_def.Cells(i, col_stage).Value
                        split_content = Split(tempTestjob, ",")
                        testjobMatched = ""
                        
                        For k = 0 To UBound(split_content)
                            If LCase("*" & testJob & "*") Like LCase("*" & split_content(k) & "*") Then
                                foundTestjob = foundTestjob And True
                                lastTestjob = tempTestjob
                                testjobMatched = split_content(k)
                                Exit For
                            End If
                        Next k
                        
                        If testjobMatched <> "" Then
                            foundTestjob = foundTestjob And True
                        Else
                            foundTestjob = foundTestjob And False
                            lastTestjob = ""
                        End If
                    Else '''Bits seq<>0
                        tempTestjob = ws_def.Cells(i, col_stage).Value
                        
                        If tempTestjob = "" Or tempTestjob = lastTestjob Then
                            split_content = Split(lastTestjob, ",")
                            testjobMatched = ""
                            
                            For k = 0 To UBound(split_content)
                                If LCase("*" & testJob & "*") Like LCase("*" & split_content(k) & "*") Then
                                    foundTestjob = foundTestjob And True
                                    testjobMatched = split_content(k)
                                    Exit For
                                End If
                            Next k
                            
                            If testjobMatched <> "" Then
                                foundTestjob = foundTestjob And True
                            Else
                                foundTestjob = foundTestjob And False
                                lastTestjob = ""
                            End If
                        Else
                            foundTestjob = foundTestjob And False
                            TheExec.Datalog.WriteComment "Row:" & i & ", column Stage sequence is incorrect. Error!!!"
                            TheExec.ErrorLogMessage "Row:" & i & ", column Stage sequence is incorrect. Error!!!"
                        End If
                    End If
                End If
                
                If foundTestjob = True Then
                    '''//If the column of "Bits" is not empty, parse the row for SELSRAM.
                    If ws_def.Cells(i, col_BitSeq).Value <> "" Then
                        currentBlock = ws_def.Cells(i, col_Block).Value
                        currentPattern = ws_def.Cells(i, col_Pattern).Value
                        bitCnt = bitCnt + 1
                        
                        '''****************************************************************************************************'''
                        '''//The area is used to check if each row follow Selsrm_Mapping_Table rules//'''
                        '''//Check if col_LogicPin has the correct content.
                        '''If Logic Pin is "PRESERVED" or "RESERVED", SRAM Pin should be "PRESERVED" or "RESERVED", too.
                        If UCase(ws_def.Cells(i, col_LogicPin).Value) = "PRESERVED" Or UCase(ws_def.Cells(i, col_LogicPin).Value) = "RESERVED" Then
                            If UCase(ws_def.Cells(i, col_LogicPin).Value) = UCase(ws_def.Cells(i, col_SramPin).Value) Then
                                foundTestjob = foundTestjob And True
                            Else
                                foundTestjob = foundTestjob And False
                                TheExec.Datalog.WriteComment "row:" & i & ", col:" & col_LogicPin & " of the sheet " & sheetName & " is incorrect. It should be listed in Vdd_Binning_Def tables. Error!!!"
                                TheExec.ErrorLogMessage "row:" & i & ", col:" & col_LogicPin & " of the sheet " & sheetName & " is incorrect. It should be listed in Vdd_Binning_Def tables. Error!!!"
                            End If
                        ElseIf bincutCorepowerPingroup <> "" Then
                            If UCase("*," & bincutCorepowerPingroup & ",*") Like UCase("*," & ws_def.Cells(i, col_LogicPin).Value & ",*") Then
                                foundTestjob = foundTestjob And True
                            Else
                                foundTestjob = foundTestjob And False
                                TheExec.Datalog.WriteComment "row:" & i & ", col:" & col_LogicPin & " of the sheet " & sheetName & " is incorrect. bincutCorepowerPingroup doesn't contain " & ws_def.Cells(i, col_LogicPin).Value & ". Error!!!"
                                TheExec.ErrorLogMessage "row:" & i & ", col:" & col_LogicPin & " of the sheet " & sheetName & " is incorrect. bincutCorepowerPingroup doesn't contain " & ws_def.Cells(i, col_LogicPin).Value & ". Error!!!"
                            End If
                        ElseIf ws_def.Cells(i, col_LogicPin).Value <> "" Then
                            foundTestjob = foundTestjob And True
                        Else
                            foundTestjob = foundTestjob And False
                        End If
                        
                        '''//Check if col_SramPin has the correct content
                        '''If Logic Pin is "PRESERVED" or "RESERVED", SRAM Pin should be "PRESERVED" or "RESERVED", too.
                        If UCase(ws_def.Cells(i, col_SramPin).Value) = "PRESERVED" Or UCase(ws_def.Cells(i, col_SramPin).Value) = "RESERVED" Then
                            If UCase(ws_def.Cells(i, col_LogicPin).Value) = UCase(ws_def.Cells(i, col_SramPin).Value) Then
                                foundTestjob = foundTestjob And True
                            Else
                                foundTestjob = foundTestjob And False
                                TheExec.Datalog.WriteComment "row:" & i & ", col:" & col_SramPin & " of the sheet " & sheetName & " is incorrect. It should be listed in Vdd_Binning_Def tables. Error!!!"
                                TheExec.ErrorLogMessage "row:" & i & ", col:" & col_SramPin & " of the sheet " & sheetName & " is incorrect. It should be listed in Vdd_Binning_Def tables. Error!!!"
                            End If
                        ElseIf bincutOtherrailPingroup <> "" Then
                            If UCase("*," & bincutOtherrailPingroup & ",*") Like UCase("*," & ws_def.Cells(i, col_SramPin).Value & ",*") Then
                                foundTestjob = foundTestjob And True
                            Else
                                foundTestjob = foundTestjob And False
                                TheExec.Datalog.WriteComment "row:" & i & ", col:" & col_SramPin & " of the sheet " & sheetName & " is incorrect. bincutOtherrailPingroup doesn't contain " & ws_def.Cells(i, col_SramPin).Value & ". Error!!!"
                                TheExec.ErrorLogMessage "row:" & i & ", col:" & col_SramPin & " of the sheet " & sheetName & " is incorrect. bincutOtherrailPingroup doesn't contain " & ws_def.Cells(i, col_SramPin).Value & ". Error!!!"
                            End If
                        ElseIf ws_def.Cells(i, col_SramPin).Value <> "" Then
                            foundTestjob = foundTestjob And True
                        Else
                            foundTestjob = foundTestjob And False
                        End If
                        
                        '''//Check if col_Selsram1 has the correct content
                        If CInt(ws_def.Cells(i, col_Selsram1).Value) = 1 Or CInt(ws_def.Cells(i, col_Selsram1).Value) = 0 Then
                            foundTestjob = foundTestjob And True
                        Else
                            foundTestjob = foundTestjob And False
                            TheExec.Datalog.WriteComment "row:" & i & ", col:" & col_Selsram1 & " of the sheet " & sheetName & " is incorrect. It should be 1 or 0. Error!!!"
                            TheExec.ErrorLogMessage "row:" & i & ", col:" & col_Selsram1 & " of the sheet " & sheetName & " is incorrect. It should be 1 or 0. Error!!!"
                        End If
                        
                        '''//Check if col_Selsram0 has the correct content
                        If CInt(ws_def.Cells(i, col_Selsram0).Value) = 1 Or CInt(ws_def.Cells(i, col_Selsram0).Value) = 0 Then
                            foundTestjob = foundTestjob And True
                        Else
                            foundTestjob = foundTestjob And False
                            TheExec.Datalog.WriteComment "row:" & i & ", col:" & col_Selsram0 & " of the sheet " & sheetName & " is incorrect. It should be 1 or 0. Error!!!"
                            TheExec.ErrorLogMessage "row:" & i & ", col:" & col_Selsram0 & " of the sheet " & sheetName & " is incorrect. It should be 1 or 0. Error!!!"
                        End If
                        
                        '''//Check if col_LogicPin is "PRESERVED".
                        '''If Logic Pin is "PRESERVED" or "RESERVED", SRAM Pin should be "PRESERVED" or "RESERVED", too.
                        If UCase(ws_def.Cells(i, col_LogicPin).Value) = "PRESERVED" Or UCase(ws_def.Cells(i, col_LogicPin).Value) = "RESERVED" Then
                            If CInt(ws_def.Cells(i, col_Selsram1).Value) = CInt(ws_def.Cells(i, col_Selsram1).Value) Then
                                foundTestjob = foundTestjob And True
                            Else
                                foundTestjob = foundTestjob And False
                                TheExec.Datalog.WriteComment "row:" & i & ", col:" & col_Selsram1 & " of the sheet " & sheetName & " is incorrect. col_Selsram1 should be same as col_Selsram0 for PRESERVED. Error!!!"
                                TheExec.ErrorLogMessage "row:" & i & ", col:" & col_Selsram1 & " of the sheet " & sheetName & " is incorrect. col_Selsram1 should be same as col_Selsram0 for PRESERVED. Error!!!"
                            End If
                        End If
                        
                        '''//Check if column stage, block, pattern are not empty when Bits sequence is 0.
                        If CInt(ws_def.Cells(i, col_BitSeq).Value) = 0 Then
                            bitCnt = 0
                        
                            If ws_def.Cells(i, col_Block).Value <> "" And ws_def.Cells(i, col_Pattern).Value <> "" Then
                                foundTestjob = foundTestjob And True
                            Else
                                foundTestjob = foundTestjob And False
                                TheExec.Datalog.WriteComment "Row:" & i & ", column Block or Pattern should not be empty when Bits sequence is 0. Error!!!"
                                TheExec.ErrorLogMessage "Row:" & i & ", column Block or Pattern should not be empty when Bits sequence is 0. Error!!!"
                            End If
                        Else
                            If bitCnt = CInt(ws_def.Cells(i, col_BitSeq).Value) Then
                                foundTestjob = foundTestjob And True
                            Else
                                foundTestjob = foundTestjob And False
                                TheExec.Datalog.WriteComment "Row:" & i & ", column Bits sequence is incorrect. Error!!!"
                                TheExec.ErrorLogMessage "Row:" & i & ", column Bits sequence is incorrect. Error!!!"
                            End If
                        End If
                    
                        '''//If "Bits"=0, start the new blockgroup.
                        If CInt(ws_def.Cells(i, col_BitSeq).Value) = 0 Then
                            blockCnt = blockCnt + 1
                            bitCnt = 0
                            
                            ReDim Preserve SelsramMapping(blockCnt)
                            
                            SelsramMapping(blockCnt).blockName = currentBlock
                            SelsramMapping(blockCnt).Pattern = currentPattern
                            
                            tempBitSeq = ws_def.Cells(i, col_BitSeq).Value
                            SelsramMapping(blockCnt).bitCount = Split(tempBitSeq, ",")
                        Else
                            If blockCnt < 0 Then
                                TheExec.ErrorLogMessage sheetName & " has the incorrect Bit sequence. Error!!!"
                                TheExec.Datalog.WriteComment sheetName & " has the incorrect Bit sequence. Error!!!"
                            Else
                                If currentBlock = "" Or currentBlock = SelsramMapping(blockCnt).blockName Then
                                    tempBitSeq = tempBitSeq & "," & ws_def.Cells(i, col_BitSeq).Value
                                    SelsramMapping(blockCnt).bitCount = Split(tempBitSeq, ",")
                                End If
                            End If
                        End If
                        
                        '''//Check logicPin with BinCut corePower pinGroup and sramPin with BinCut otherRail pinGroup.
                        '''//Check "Logic Pins"
                        If bitCnt = 0 Then
                            '''//Check "Logic Pins"
                            tempLogicPinSeq = ws_def.Cells(i, col_LogicPin).Value
                            tempAlphaSeq = ws_def.Cells(i, col_Alpha).Value
                            
                            '''//Check SramPin
                            tempSramPinSeq = ws_def.Cells(i, col_SramPin).Value
                            
                            '''Check Selsram1
                            tempSelsram1Seq = ws_def.Cells(i, col_Selsram1).Value
                            
                            '''Check Selsram0
                            tempSelsram0Seq = ws_def.Cells(i, col_Selsram0).Value
                            
                            '''Check Comment
                            tempCommentSeq = ws_def.Cells(i, col_comment).Value
                        Else
                            '''//Check "Logic Pins"
                            tempLogicPinSeq = tempLogicPinSeq & "," & ws_def.Cells(i, col_LogicPin).Value
                            tempAlphaSeq = tempAlphaSeq & "," & ws_def.Cells(i, col_Alpha).Value
                            
                            '''//Check SramPin
                            tempSramPinSeq = tempSramPinSeq & "," & ws_def.Cells(i, col_SramPin).Value
                            
                            '''Check Selsram1
                            tempSelsram1Seq = tempSelsram1Seq & "," & ws_def.Cells(i, col_Selsram1).Value
                            
                            '''Check Selsram0
                            tempSelsram0Seq = tempSelsram0Seq & "," & ws_def.Cells(i, col_Selsram0).Value
                            
                            '''Check Comment
                            tempCommentSeq = tempCommentSeq & "," & ws_def.Cells(i, col_comment).Value
                        End If
                        
                        SelsramMapping(blockCnt).logic_Pin = Split(tempLogicPinSeq, ",")
                        SelsramMapping(blockCnt).alpha = Split(tempAlphaSeq, ",")
                        SelsramMapping(blockCnt).sram_Pin = Split(tempSramPinSeq, ",")
                        SelsramMapping(blockCnt).SelSrm1 = Split(tempSelsram1Seq, ",")
                        SelsramMapping(blockCnt).SelSrm0 = Split(tempSelsram0Seq, ",")
                        SelsramMapping(blockCnt).comment = Split(tempCommentSeq, ",")
                        
                        ''''//selsramLogicPin
                        If LCase("*," & selsramLogicPin & ",*") Like LCase("*," & ws_def.Cells(i, col_LogicPin).Value & ",*") Then
                            '''Do nothing
                        Else
                            '''If Logic Pin is "PRESERVED" or "RESERVED", SRAM Pin should be "PRESERVED" or "RESERVED", too.
                            If UCase(ws_def.Cells(i, col_LogicPin).Value) <> UCase("PRESERVED") Or UCase(ws_def.Cells(i, col_LogicPin).Value) <> UCase("RESERVED") Then
                                If selsramLogicPin <> "" Then
                                    selsramLogicPin = selsramLogicPin & "," & UCase(ws_def.Cells(i, col_LogicPin).Value)
                                    tempPinalpha = tempPinalpha & "," & UCase(ws_def.Cells(i, col_Alpha).Value)
                                Else
                                    selsramLogicPin = UCase(ws_def.Cells(i, col_LogicPin).Value)
                                    tempPinalpha = UCase(ws_def.Cells(i, col_Alpha).Value)
                                End If
                            End If
                        End If
                        
                        ''''//selsramSramPin
                        If LCase("*," & selsramSramPin & ",*") Like LCase("*," & ws_def.Cells(i, col_SramPin).Value & ",*") Then
                            '''Do nothing
                        Else
                            '''If Logic Pin is "PRESERVED" or "RESERVED", SRAM Pin should be "PRESERVED" or "RESERVED", too.
                            If UCase(ws_def.Cells(i, col_SramPin).Value) <> UCase("PRESERVED") Or UCase(ws_def.Cells(i, col_SramPin).Value) <> UCase("RESERVED") Then
                                If selsramSramPin <> "" Then
                                    selsramSramPin = selsramSramPin & "," & UCase(ws_def.Cells(i, col_SramPin).Value)
                                Else
                                    selsramSramPin = UCase(ws_def.Cells(i, col_SramPin).Value)
                                End If
                            End If
                        End If
                    End If '''If ws_def.Cells(i, col_BitSeq).value <> "" Then
                End If '''If foundTestjob = True Then
            Next i '''For i = Row_of_Title + 1 To MaxRow
        Else
            TheExec.Datalog.WriteComment sheetName & " doesn't have the correct header for Parsing_SELSRM_Mapping_Table. Error!!!"
            TheExec.ErrorLogMessage sheetName & " doesn't have the correct header for Parsing_SELSRM_Mapping_Table. Error!!!"
        End If
    End If '''If isSheetFound = True
    
    '''//Check if the pinGroup is empty or not.
'''ToDo: Maybe we can create the dictionary for selsramPin when parsing the table "SELSRM_Mapping_Table"...
    If selsramLogicPin <> "" And selsramLogicPinalpha <> "" Or selsramSramPin <> "" Then
        selsramPin = selsramLogicPin
        
        '''//Alphabet of selsramLogicPin
        selsramLogicPinalpha = tempPinalpha
        
        '''//Add selsramSramPin into selsramPin
        selsramPin = selsramPin & "," & selsramSramPin
        
        '''//Generate selsram PinGroups
        selsramLogicPingroup = Split(selsramLogicPin, ",")
        selsramLogicPinalphagroup = Split(selsramLogicPinalpha, ",")
        selsramSramPingroup = Split(selsramSramPin, ",")
        selsramPingroup = Split(selsramPin, ",")
        
        '''*********************************************************************************************************************************************'''
        '''//For conventional projects that have all BinCut powerPins with pattern vbump, we can override selsramPin by FullBinCutPowerinFlowSheet here.
        '''//For projects that only have certain Selsram powerPins with pattern vbump, no need to change vbt code.
        '''If Flag_SyncUp_DCVS_Output_enable = False and Flag_Enable_Rail_Switch=true, only selsram powerpins apply payload voltages to DCVS Valt.
        '''<org>
            'If Flag_SyncUp_DCVS_Output_enable = True Then
            '    selsramPin = FullBinCutPowerinFlowSheet
            'End If
        '''<new>
        '''20200210: As per discussion with C651 Chris and Toby, they will provide patterns with vbump for all BinCut powerDomains.
        selsramPin = FullBinCutPowerinFlowSheet
        '''*********************************************************************************************************************************************'''
        
        '''//Once if the table parsing is done, set the flag as true for GradeSearch instances.
        Flag_SelsrmMappingTable_Parsed = True
    Else
        TheExec.Datalog.WriteComment sheetName & " doesn't contain correct LogicPin or SramPin or Alphabet for SELSRAM. Error!!!"
        TheExec.ErrorLogMessage sheetName & " doesn't contain correct LogicPin or SramPin or Alphabet for SELSRAM. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Parsing_SELSRM_Mapping_Table"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Parsing_SELSRM_Mapping_Table"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210908: Modified to merge the vbt code of the vbt function HVCC_OtherRail_Pwr_Setting into the vbt function OtherRail_Pwr_Setting, as discussed with TSMC ZYLINI.
'20210906: Modified to remove the redundant argument "ByRef voltage_SelsrmBitCalc As SiteDouble".
'20201116: Modified to use "testCondition = LCase(Trim(Get_BinCut_TestCondition()))".
'20201029: Modified to use "Public Type Instance_Info".
'20201029: Modified to replace the argument "other_rail_mode as Integer" with "idxPin_OtherRail As Integer".
'20200806: Modified to use siteDouble for Calculate_LVCC_for_CorePower, Calculate_HVCC_for_CorePower, and Calculate_Voltage_for_OtherRail.
'20200803: Modified to use the argument "ByRef voltage_SelsrmBitCalc As SiteDouble".
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20200711: Modified to change the argumens types and store BinCut payload voltages into the siteDouble array.
'20200710: Modified to merge the redundant site-loop for voltage calculation.
'20200430: Modified to print siteNumber for the site with the incorrect voltage.
'20200429: Modified to use "Get_BinCut_TestCondition" to get BinCut testCondition.
'20200106: Modified to remove the ErrorLogMessage.
'20191227: Modified to check the string of testCondition in Non_Binning_Rail.
'20191202: Modified for the revised initVddBinCondition.
'20191127: Modified for the revised InitVddBinTable.
'20190507: Modified to add "Cdec" to avoid double format accuracy issues.
'20190319: Modified to call the same function "Calculate_Voltage_for_OtherRail" for voltages calculation of OtherRail_Pwr_Setting and HVCC_OtherRail_Pwr_Setting.
'20190314: Modified for calling the function Calculate_LVCC_for_OtherRail to calculate CorePower voltage.
'20190307: Modified for compatible with conventional projects(SRAM with CorePower Pmode).
'20190304: Modified for calling the new funciton Calculate_Extra_Voltage_for_PowerRail.
'20181004: Modified to change testCondition for postBinCut and retention.
'20180913: Modified to prevent output values from DCVS out of range.
Public Function OtherRail_Pwr_Setting(inst_info As Instance_Info, idxPin_OtherRail As Integer, passBinCut As SiteLong, ByRef voltage_OtherRail As SiteDouble)
    Dim site As Variant
    Dim str_testCondition As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''<Keyword replacement of BinCut test condition>
'''20180926: Currently C651 Toby didn't define "bin result" or "product-*gb" for postBinCut and retention tests.
'''So that we define the voltage as "VBIN_RESULT(P_mode).Grade". If we get the definition from Toby, we will update this.
'''<Warning!!!>
'''ToDo: Please discuss this with C651 project DRIs to see if we can use the keyword in the instance names to decide the keyword replacement of BinCut test condition.
'''//==================================================================================================================================================================================//'''
    For Each site In TheExec.sites
        '''//Get BinCut testCondition for the selected powerDomain(OtherRail).
        str_testCondition = LCase(Trim(Get_BinCut_TestCondition(inst_info, idxPin_OtherRail, passBinCut(site))))
        
        '''//Calculate BinCut payload voltage of OtherRail according to testCondtion from sheet "Non_Binning_Rail".
        Call Calculate_Voltage_for_OtherRail(inst_info, idxPin_OtherRail, passBinCut, str_testCondition, voltage_OtherRail)
        
        '''//Floor BinCut payload voltage for BV test instances.
        If CDec(voltage_OtherRail) > 0 Then '''prevent DCVS from out of range error.
            '''20210908: Modified to merge the vbt code of the vbt function HVCC_OtherRail_Pwr_Setting into the vbt function OtherRail_Pwr_Setting, as discussed with TSMC ZYLINI.
            If LCase(inst_info.inst_name) Like "*_hbv" Then '''HBV
                '''//HBV instances ceiling the value of BinCut payload voltage.
                voltage_OtherRail(site) = Ceiling(voltage_OtherRail(site)) '''unit: mV
            ElseIf LCase(inst_info.inst_name) Like "*_bv" Then '''BV
                '''//BV instances floor the value of BinCut payload voltage.
                voltage_OtherRail(site) = Floor(voltage_OtherRail(site)) '''unit: mV
            Else
                TheExec.Datalog.WriteComment "instance:" & inst_info.inst_name & ", it doesn't contain any keyword about _bv or _hbv for OtherRail_Pwr_Setting. Error!!!"
                TheExec.ErrorLogMessage "instance:" & inst_info.inst_name & ", it doesn't contain any keyword about _bv or _hbv for OtherRail_Pwr_Setting. Error!!!"
            End If
        Else
            voltage_OtherRail(site) = 0
            TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ", BinCut payload voltage should be larger than 0. Error occurs in OtherRail_Pwr_Setting while " & inst_info.inst_name & ". Error!!!"
            'TheExec.ErrorLogMessage "site:" & site & "," & VddBinName(inst_info.p_mode) & ",powerDomain:" & ", BinCut payload voltage should be larger than 0. Error occurs in OtherRail_Pwr_Setting while " & inst_info.inst_name & ". Error!!!"
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of OtherRail_Pwr_Setting"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210806: Modified to remove the redundant vbt function Get_SRAM_VTH.
'20200703: Modiifed to use "check_Sheet_Range".
'20200324: Modified to use the data types "PWRBIN_SPEC_Type", "PWRBIN_CONDITION_Type".
'20190426: Modified to use the function "Find_Sheet".
'20190312: Modified to Add for New Power_binning_control_table.
'20180913: Modified for Multiple Power Binning.
Public Function Parsing_Power_Bin_Table(sheetName As String)
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim row As Integer, col As Integer
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim headerFound As Boolean
    Dim row_of_title As Long
    Dim col_Seq As Long
    Dim col_Bin1 As Long
    Dim col_BinX As Long
    Dim col_BinY As Long
    Dim idxCurrentCondition As Integer
    Dim strSheetBin1 As String
    Dim strSheetBinX As String
    Dim strSheetBinY As String
    Dim strSheetBinAll As String
    Dim strAry_sheetname() As String
    Dim idxSheet As Integer
    Dim isSheetFound As Boolean
On Error GoTo errHandler
    '''*****************************************************************'''
    '''//Check if the sheet exists
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        idxCurrentCondition = -1
        headerFound = False
        strSheetBin1 = ""
        strSheetBinX = ""
        strSheetBinY = ""
        strSheetBinAll = ""
        PwrBin_SheetCnt = -1
        
        '''//Check the header of the table
        For row = 1 To MaxRow
            For col = 1 To maxcol
                If UCase(ws_def.Cells(row, col)) = "SEQUENCE" Then
                    col_Seq = col
                    headerFound = True
                ElseIf UCase(ws_def.Cells(row, col)) = "BIN1" Then
                    col_Bin1 = col
                ElseIf UCase(ws_def.Cells(row, col)) = "BINX" Then
                    col_BinX = col
                ElseIf UCase(ws_def.Cells(row, col)) = "BINY" Then
                    col_BinY = col
                End If
            Next col
    
            If headerFound = True Then
                row_of_title = row
                Exit For
            End If
        Next row
        
        If headerFound = True Then
            '''//column "Bin1"
            idxCurrentCondition = idxCurrentCondition + 1
            ReDim Preserve AllPwrBin(idxCurrentCondition)
            AllPwrBin(idxCurrentCondition).passBinCut = 1
            AllPwrBin(idxCurrentCondition).harvestUsed = False
            AllPwrBin(idxCurrentCondition).harvestBin = 0
            
            '''//column "BinX"
            If col_BinX > 0 Then
                If col_Bin1 > 0 Then
                    idxCurrentCondition = idxCurrentCondition + 1
                    ReDim Preserve AllPwrBin(idxCurrentCondition)
                    AllPwrBin(idxCurrentCondition).passBinCut = 2
                    AllPwrBin(idxCurrentCondition).harvestUsed = False
                    AllPwrBin(idxCurrentCondition).harvestBin = 0
                Else
                    headerFound = False
                    TheExec.Datalog.WriteComment "The header of " & sheetName & " is incorrect due to column BinX. Error!!!"
                    TheExec.ErrorLogMessage "The header of " & sheetName & " is incorrect due to column BinX. Error!!!"
                End If
            End If
            
            '''//column "BinY"
            If col_BinY > 0 Then
                If col_Bin1 > 0 And col_BinX > 0 Then
                    idxCurrentCondition = idxCurrentCondition + 1
                    ReDim Preserve AllPwrBin(idxCurrentCondition)
                    AllPwrBin(idxCurrentCondition).passBinCut = 3
                    AllPwrBin(idxCurrentCondition).harvestUsed = False
                    AllPwrBin(idxCurrentCondition).harvestBin = 0
                Else
                    headerFound = False
                    TheExec.Datalog.WriteComment "The header of " & sheetName & " is incorrect due to column BinY. Error!!!"
                    TheExec.ErrorLogMessage "The header of " & sheetName & " is incorrect due to column BinY. Error!!!"
                End If
            End If
        Else
            TheExec.Datalog.WriteComment "The header of " & sheetName & " is incorrect. Error!!!"
            TheExec.ErrorLogMessage "The header of " & sheetName & " is incorrect. Error!!!"
        End If
    
        '''//Parsing all power binning sheets
        If headerFound = True Then
            row = row_of_title + 1
            
            While LCase(ws_def.Cells(row, col_Seq).Value) <> "" And LCase(ws_def.Cells(row, col_Seq).Value) <> "end"
                If ws_def.Cells(row, col_Bin1).Value <> "" Then
                    If strSheetBin1 <> "" Then
                        strSheetBin1 = strSheetBin1 & "," & ws_def.Cells(row, col_Bin1).Value
                    Else
                        strSheetBin1 = ws_def.Cells(row, col_Bin1).Value
                    End If
                End If
                
                If col_BinX > 0 Then
                    If ws_def.Cells(row, col_BinX).Value <> "" Then
                        If strSheetBinX <> "" Then
                            strSheetBinX = strSheetBinX & "," & ws_def.Cells(row, col_BinX).Value
                        Else
                            strSheetBinX = ws_def.Cells(row, col_BinX).Value
                        End If
                    End If
                End If
                
                If col_BinY > 0 Then
                    If ws_def.Cells(row, col_BinY).Value <> "" Then
                        If strSheetBinY <> "" Then
                            strSheetBinY = strSheetBinX & "," & ws_def.Cells(row, col_BinY).Value
                        Else
                            strSheetBinY = ws_def.Cells(row, col_BinY).Value
                        End If
                    End If
                End If
                
                row = row + 1
            Wend
            
            If strSheetBin1 <> "" Then
                strSheetBinAll = strSheetBin1
            
                If strSheetBinX <> "" Then
                    strSheetBinAll = strSheetBinAll & "," & strSheetBinX
                End If
                
                If strSheetBinY <> "" Then
                    strSheetBinAll = strSheetBinAll & "," & strSheetBinY
                End If
            Else
                headerFound = False
                TheExec.Datalog.WriteComment "No correct sheetName for Bin1 exists in " & sheetName & ". Error!!!"
                TheExec.ErrorLogMessage "No correct sheetName for Bin1 exists in " & sheetName & ". Error!!!"
            End If
        End If
        
        If headerFound = True Then
            strAry_sheetname = Split(strSheetBinAll, ",")
            PwrBin_SheetCnt = UBound(strAry_sheetname) + 1
            
            ReDim AllPwrBin(0).TestSpec(0)
            ReDim AllPwrBin(0).TestSpec(0).specUsed(UBound(strAry_sheetname))
            ReDim AllPwrBin(0).TestSpec(0).specCustomized(UBound(strAry_sheetname))
            
            If col_BinX > 0 Then
                ReDim AllPwrBin(1).TestSpec(0)
                ReDim AllPwrBin(1).TestSpec(0).specUsed(UBound(strAry_sheetname))
                ReDim AllPwrBin(1).TestSpec(0).specCustomized(UBound(strAry_sheetname))
            End If
            
            If col_BinY > 0 Then
                ReDim AllPwrBin(2).TestSpec(0)
                ReDim AllPwrBin(2).TestSpec(0).specUsed(UBound(strAry_sheetname))
                ReDim AllPwrBin(2).TestSpec(0).specCustomized(UBound(strAry_sheetname))
            End If
            
            For idxSheet = 0 To UBound(strAry_sheetname)
                '''Parse the sheet
                PwrBin_SheetnameDict.Add strAry_sheetname(idxSheet), idxSheet
                ReDim Preserve PwrBin_Sheet(idxSheet)
                Call Parsing_Row_Column(strAry_sheetname(idxSheet), idxSheet)
                Flag_PowerBinningTable_Parsed = True
                
                If LCase(",*" & strSheetBin1 & ",*") Like LCase(",*" & strAry_sheetname(idxSheet) & ",*") Then
                    AllPwrBin(0).TestSpec(0).specUsed(idxSheet) = True
                    AllPwrBin(0).TestSpec(0).specCustomized(idxSheet) = PwrBin_Sheet(idxSheet).spec
                Else
                    AllPwrBin(0).TestSpec(0).specUsed(idxSheet) = False
                End If
                
                If col_BinX > 0 Then
                    If LCase(",*" & strSheetBinX & ",*") Like LCase(",*" & strAry_sheetname(idxSheet) & ",*") Then
                        AllPwrBin(1).TestSpec(0).specUsed(idxSheet) = True
                        AllPwrBin(1).TestSpec(0).specCustomized(idxSheet) = PwrBin_Sheet(idxSheet).spec
                    Else
                        AllPwrBin(1).TestSpec(0).specUsed(idxSheet) = False
                    End If
                End If
                
                If col_BinY > 0 Then
                    If LCase(",*" & strSheetBinY & ",*") Like LCase(",*" & strAry_sheetname(idxSheet) & ",*") Then
                        AllPwrBin(2).TestSpec(0).specUsed(idxSheet) = True
                        AllPwrBin(2).TestSpec(0).specCustomized(idxSheet) = PwrBin_Sheet(idxSheet).spec
                    Else
                        AllPwrBin(2).TestSpec(0).specUsed(idxSheet) = False
                    End If
                End If
            Next idxSheet
        End If
    End If '''If isSheetFound = True
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Parsing_Power_Bin_Table"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Parsing_Power_Bin_Table"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210705: Modified to check "power_binning" in dict_EfuseCategory2BinCutTestJob.
'20210628: Modified to use gb_str_EfuseCategory_for_powerbinning.
'20210518: Modified to skip checking if currentTestname <> previousTestname.
'20210113: Modified to check col_fusePwrbin > 0, col_fuseValue>0, LCase(currentHarvestbin) = "2", and idxSheet <= -1.
'20200703: Modiifed to use "check_Sheet_Range".
'20200528: Modified to check header of the table.
'20200305: Modified to reduce the flow complexity (requested by C651 Toby).
'20200302: Modified to store columns "PASS: power_binning" and "PASS: fuse_name2".
'20200227: Modified to store "fuse_value" for SPEC.
'20200226: Modified to Check if next spec exists for the previous spec.
'20200218: Created to parse the new format of PowerBinning table for Harvest.
Public Function Parsing_Power_Bin_Table_Harvest(sheetName As String)
    Dim wb As Workbook
    Dim ws_def As Worksheet
    '''check the sheet
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim headerFound As Boolean
    Dim row As Integer, col As Integer
    '''columns
    Dim row_of_title As Integer
    Dim col_passbin As Integer
    Dim col_harvestbin As Integer
    Dim col_fusePwrbin As Integer
    Dim col_fusevalue As Integer
    Dim col_specname As Integer
    Dim col_pwrbin_sheetname() As Integer
    Dim strAry_sheetname() As String
    Dim binNumber As Integer
    Dim idxSheet As Integer
    '''variables
    Dim i As Integer, j As Integer
    Dim previousPassbin As Double
    Dim currentPassbin As Double
    Dim previousHarvestbin As String
    Dim currentHarvestbin As String
    Dim idxSpec As Integer
    Dim foundTestName As Boolean
    Dim previousTestname As String
    Dim currentTestname As String
    Dim idxPreviousCondition As Integer
    Dim idxCurrentCondition As Integer
    Dim cntSpec As Integer
    Dim isSheetFound As Boolean
    Dim split_content() As String
On Error GoTo errHandler
    '''*****************************************************************'''
    '''//Check if the sheet exists
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        headerFound = False
        idxSheet = -1
        idxSpec = -1
        previousPassbin = 0
        currentPassbin = 0
        previousHarvestbin = ""
        currentHarvestbin = ""
        foundTestName = True
        PwrBin_SheetCnt = -1
        previousTestname = ""
        currentTestname = ""
        idxPreviousCondition = -1
        idxCurrentCondition = -1
        cntSpec = 0
        
        '''//Check the header of the table
        '''Header: "INPUT: product_identifier", "INPUT: harvesting_bin", "PwrBin1_MB1_V2" "PwrBin1_MB2_V2", "PwrBin2_MB1_V2", "PwrBin2_MB2_V2", "PASS: power_binning", "PASS: fuse_name2", "FAIL", "comment".
        For row = 1 To MaxRow
            For col = 1 To maxcol
                If UCase(ws_def.Cells(row, col).Value) Like UCase("INPUT: product_identifier") Then '''Product_identifier
                    col_passbin = col
                    row_of_title = row
                End If
                
                If row_of_title > 0 Then
                    If UCase(ws_def.Cells(row_of_title, col).Value) Like UCase("INPUT: harvesting_bin") Or UCase(ws_def.Cells(row_of_title, col).Value) Like UCase("INPUT: gfx_harvesting_bin") Then '''Harvest_bin
                        col_harvestbin = col
                    ElseIf UCase(ws_def.Cells(row_of_title, col).Value) Like UCase("PASS: power_binning") Then '''PASS: power_binning
                        col_fusePwrbin = col
                        '''//Check if "PASS: power_binning" exists in the header of PowerBinning table.
                        '''20210705: Modified to check "power_binning" in dict_EfuseCategory2BinCutTestJob.
                        split_content = Split(UCase(ws_def.Cells(row_of_title, col).Value), "PASS:")
                        
                        If dict_EfuseCategory2BinCutTestJob.Exists(UCase(Trim(split_content(UBound(split_content))))) = True Then
                            gb_str_EfuseCategory_for_powerbinning = UCase(Trim(split_content(UBound(split_content))))
                        Else
                            gb_str_EfuseCategory_for_powerbinning = ""
                            'TheExec.Datalog.WriteComment "sheet:" & sheetName & ",header:" & ws_def.Cells(row_of_title, col).Value & ", it isn't defined in Efuse_BitDef_Table. Please check Efuse_BitDef_Table for this. Error!!!"
                            'TheExec.ErrorLogMessage "sheet:" & sheetName & ",header:" & ws_def.Cells(row_of_title, col).Value & ", it isn't defined in Efuse_BitDef_Table. Please check Efuse_BitDef_Table for this. Error!!!"
                        End If
                    ElseIf UCase(ws_def.Cells(row_of_title, col).Value) Like UCase("PASS: fuse_name2") Then '''PASS: fuse_name2
                        col_fusevalue = col
                    ElseIf UCase(ws_def.Cells(row_of_title, col).Value) Like UCase("Comment") Then '''Spec_name
                        col_specname = col
                    ElseIf UCase(ws_def.Cells(row_of_title, col).Value) Like UCase("PwrScreen*") Then '''Spec sheetname
                        If Find_Sheet(ws_def.Cells(row_of_title, col).Value) = True Then
                            idxSheet = idxSheet + 1
                            ReDim Preserve col_pwrbin_sheetname(idxSheet)
                            ReDim Preserve strAry_sheetname(idxSheet)
                            col_pwrbin_sheetname(idxSheet) = col
                            strAry_sheetname(idxSheet) = ws_def.Cells(row_of_title, col).Value
                            
                            '''//Parsing all power binning sheets as bin1 sheets
                            binNumber = 1
                            
                            '''Parse the sheet
                            PwrBin_SheetnameDict.Add strAry_sheetname(idxSheet), idxSheet
                            ReDim Preserve PwrBin_Sheet(idxSheet)
                            Call Parsing_Row_Column(strAry_sheetname(idxSheet), idxSheet)
                            Flag_PowerBinningTable_Parsed = True
                        Else
                            TheExec.Datalog.WriteComment "PowerBinning sheet:" & ws_def.Cells(row_of_title, col).Value & ", it doesn't exist in the workbook. Error!!!"
                            TheExec.ErrorLogMessage "PowerBinning sheet:" & ws_def.Cells(row_of_title, col).Value & ", it doesn't exist in the workbook. Error!!!"
                        End If
                    End If
                End If
            Next col
            
            If row_of_title > 0 Then
                If col_passbin > 0 And col_harvestbin > 0 And col_specname > 0 And idxSheet > -1 Then
                    headerFound = True
                    PwrBin_SheetCnt = idxSheet + 1
                    Flag_Enable_PowerBinning_Harvest = True
                Else
                    headerFound = False
                    
                    If col_passbin = 0 Then
                        TheExec.Datalog.WriteComment "Column product_identifier doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column product_identifier doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_harvestbin = 0 Then
                        TheExec.Datalog.WriteComment "Column harvesting_bin doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column harvesting_bin doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    '''20210113: Modified to check col_fusePwrbin > 0, col_fuseValue>0, LCase(currentHarvestbin) = "2", and idxSheet <= -1.
                    If col_specname = 0 Then
                        TheExec.Datalog.WriteComment "Column Comment doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column Comment doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If idxSheet <= -1 Then
                        TheExec.Datalog.WriteComment "Header of " & sheetName & " doesn't contain any powerbinning sheetname with keyword PwrScreen. Error!!!"
                        TheExec.ErrorLogMessage "Header of " & sheetName & " doesn't contain any powerbinning sheetname with keyword PwrScreen. Error!!!"
                    End If
                End If
                
                Exit For
            End If
        Next row
        
        '''//Parsing the spec for "Product_identifier" and "Harvest_bin".
        If headerFound = True Then
            row = row_of_title + 1
            
            While LCase(ws_def.Cells(row, col_passbin).Value) <> "" And LCase(ws_def.Cells(row, col_passbin).Value) <> "end"
                currentPassbin = CInt(ws_def.Cells(row, col_passbin).Value) + 1 '''Passbin=product_identifier+1
                currentHarvestbin = ws_def.Cells(row, col_harvestbin).Value
                currentTestname = ws_def.Cells(row, col_specname).Value
                
                If currentHarvestbin <> "" And Flag_Enable_PowerBinning_Harvest = False Then
                    Flag_Enable_PowerBinning_Harvest = True
                End If
                
                '''20210518: Modified to skip checking if currentTestname <> previousTestname.
                If currentPassbin <= MaxPassBinCut Then 'And currentTestname <> previousTestname Then
                    currentHarvestbin = ws_def.Cells(row, col_harvestbin).Value
                    
                    '''//Check if current Harvest_bin is same as previous Harvest_bin.
                    If LCase(currentHarvestbin) = "1" Or LCase(currentHarvestbin) = "0" _
                    Or LCase(currentHarvestbin) = "x" Or LCase(currentHarvestbin) = "" Then
                        foundTestName = True
                    Else
                        foundTestName = False
                        TheExec.Datalog.WriteComment "PowerBinning Harvest sequence: " & sheetName & " has the incorrect Fuse_Value in row " & row & ". Error!!!"
                        TheExec.ErrorLogMessage "PowerBinning Harvest sequence: " & sheetName & " has the incorrect Fuse_Value in row " & row & ". Error!!!"
                    End If
                    
                    '''//Check if Two specs with same passbin and same harvest_bin use the same spec sheets.
                    If foundTestName = True Then
                        If (currentPassbin = previousPassbin) And (currentHarvestbin = previousHarvestbin) Then
                            For i = 0 To idxSheet
                                If (ws_def.Cells(row, col_pwrbin_sheetname(i)).Value <> "" And ws_def.Cells(row - 1, col_pwrbin_sheetname(i)).Value <> "") _
                                Or (ws_def.Cells(row, col_pwrbin_sheetname(i)).Value = "" And ws_def.Cells(row - 1, col_pwrbin_sheetname(i)).Value = "") Then
                                    foundTestName = True
                                Else
                                    foundTestName = False
                                    TheExec.Datalog.WriteComment "PowerBinning Harvest sequence: " & sheetName & " doesn't have the incorrect format in row " & row & ". Error!!!"
                                    TheExec.ErrorLogMessage "PowerBinning Harvest sequence: " & sheetName & " doesn't have the incorrect format in row " & row & ". Error!!!"
                                End If
                            Next i
                        End If
                    End If
                    
                    '''//Format is correct, so that parse the row and store the powerbinning table into allpwrbin array.
                    If foundTestName = True Then
                        If (currentPassbin = previousPassbin) And (currentHarvestbin = previousHarvestbin) Then
                            idxCurrentCondition = idxCurrentCondition
                            idxSpec = idxSpec + 1
                        Else
                            idxCurrentCondition = idxCurrentCondition + 1
                            idxSpec = 0
                        End If
                    
                        If idxCurrentCondition <> idxPreviousCondition Then
                            ReDim Preserve AllPwrBin(idxCurrentCondition)
                        End If
                        
                        ReDim Preserve AllPwrBin(idxCurrentCondition).TestSpec(idxSpec)
                        
                        '''//Passbin number ("Product_identifier")
                        AllPwrBin(idxCurrentCondition).passBinCut = currentPassbin
                        
                        '''//Check if Harvest bin is used
                        '''//If AllPwrBin(currentPassbin).harvestUsed = False, AllPwrBin(currentPassbin).condition(0)
                        If LCase(currentHarvestbin) = "x" Or LCase(currentHarvestbin) = "" Then
                            AllPwrBin(idxCurrentCondition).harvestUsed = False
                            AllPwrBin(idxCurrentCondition).harvestBin = 0
                        ElseIf LCase(currentHarvestbin) = "1" Or LCase(currentHarvestbin) = "0" Then
                            AllPwrBin(idxCurrentCondition).harvestUsed = True
                            AllPwrBin(idxCurrentCondition).harvestBin = CInt(currentHarvestbin)
                        End If
                        
                        cntSpec = cntSpec + 1
                        '''
                        If PwrBin_SpecIdx2SpecNameDict.Exists(cntSpec) Then
                            '''Do nothing
                        Else
                            '''20210113: Modified to check col_fusePwrbin > 0, col_fuseValue>0, LCase(currentHarvestbin) = "2" and idxSheet <= -1.
                            If col_fusevalue > 0 Then
                                PwrBin_SpecIdx2SpecNameDict.Add cntSpec, (ws_def.Cells(row, col_fusePwrbin).Value & "," & ws_def.Cells(row, col_fusevalue).Value & "," & ws_def.Cells(row, col_specname).Value)
                            Else
                                PwrBin_SpecIdx2SpecNameDict.Add cntSpec, "," & "," & ws_def.Cells(row, col_specname).Value
                            End If
                        End If
                        
                        AllPwrBin(idxCurrentCondition).TestSpec(idxSpec).idxAllSpec = cntSpec
                        
                        '''//TestName
                        AllPwrBin(idxCurrentCondition).TestSpec(idxSpec).testName = currentTestname
                        
                        '''//"PASS: power_binning"
                        '''20210113: Modified to check col_fusePwrbin > 0, col_fuseValue>0, LCase(currentHarvestbin) = "2", and idxSheet <= -1.
                        If col_fusePwrbin > 0 Then
                            AllPwrBin(idxCurrentCondition).TestSpec(idxSpec).fusePwrbin = ws_def.Cells(row, col_fusePwrbin).Value
                        End If
                        
                        '''//"PASS: fuse_name2"
                        '''20210113: Modified to check col_fusePwrbin > 0, col_fuseValue>0, LCase(currentHarvestbin) = "2" and idxSheet <= -1.
                        If col_fusevalue > 0 Then
                            AllPwrBin(idxCurrentCondition).TestSpec(idxSpec).fuseValue = ws_def.Cells(row, col_fusevalue).Value
                        End If
                        
                        '''//Check if next spec exists for the previous spec "AllPwrBin(idxSpec - 1)".
                        AllPwrBin(idxCurrentCondition).TestSpec(idxSpec).haveNextSpec = False
                        
                        If idxSpec > 0 Then
                            AllPwrBin(idxCurrentCondition).TestSpec(idxSpec - 1).haveNextSpec = True
                        End If
                        
                        For i = 0 To idxSheet
                            ReDim Preserve AllPwrBin(idxCurrentCondition).TestSpec(idxSpec).specUsed(i)
                            ReDim Preserve AllPwrBin(idxCurrentCondition).TestSpec(idxSpec).specCustomized(i)
                            
                            If ws_def.Cells(row, col_pwrbin_sheetname(i)).Value <> "" Then
                                AllPwrBin(idxCurrentCondition).TestSpec(idxSpec).specUsed(i) = True
                                AllPwrBin(idxCurrentCondition).TestSpec(idxSpec).specCustomized(i) = CDbl(ws_def.Cells(row, col_pwrbin_sheetname(i)).Value)
                            Else
                                AllPwrBin(idxCurrentCondition).TestSpec(idxSpec).specUsed(i) = False
                            End If
                        Next i
                    End If
                    
                    previousTestname = currentTestname
                    idxPreviousCondition = idxCurrentCondition
                    previousPassbin = currentPassbin
                    previousHarvestbin = currentHarvestbin
                Else
                    TheExec.Datalog.WriteComment "Sheet: " & sheetName & " has the incorrect product_identifier or the incorrect TestName in row " & row & ". Error!!!"
                    TheExec.ErrorLogMessage "Sheet: " & sheetName & " has the incorrect product_identifier or the incorrect TestName in row " & row & ". Error!!!"
                End If
                
                row = row + 1
            Wend
        Else
            TheExec.Datalog.WriteComment sheetName & " doesn't have the correct header for Parsing_Power_Bin_Table_Harvest. Error!!!"
            TheExec.ErrorLogMessage sheetName & " doesn't have the correct header for Parsing_Power_Bin_Table_Harvest. Error!!!"
        End If
    End If '''If isSheetFound = True
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Parsing_Power_Bin_Table_Harvest"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Parsing_Power_Bin_Table_Harvest"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201211: Modified to parse the powerbinning table without Other_Mode, requested by C651 Toby and Ben.
'20201211: Modified to parse new format "Binned_Mode","Bin Voltage (mV)","IDS (mA)","Vdd0 (mV)","Vdd1 (mV)" of the header, requested by C651 Toby and Ben.
'20201111: Modified to set Binned_Mode-loop and Other_Mode-loop start from 0.
'20201110: Modified to check cell contents of "Bin Voltage" and "IDS" for Binned_Mode and Other_Mode in Power_Binning table.
'20201110: Modified to check if any duplicate Binned_Mode or Other_Mode exists, requested by PCLINZG and SYLIUU.
'20201110: Modified to revised the vbt code for parsing PowerBinning sheets with the new format.
'20200831: Modified to define the array size of PwrBin_Sheet.
'20200703: Modiifed to use "check_Sheet_Range".
'20200608: Modified to check columns of "Bin Voltage" and "IDS".
'20200528: Modified to check header of the table.
'20200324: Modified to merge "Parsing_Row_Column_Harvest" into this function.
'20200220: Modified to parsing PowerBinning sheets for Harvest.
'20190813: Modified the parsing method for power binning tables.
'20190812: Since Sicily otherRail VDD_CPU_SRAM liked to VDD_PCPU p_mode, we modified the vbt code.
'20190426: Modified to use the function "Find_Sheet".
'20190308: Modified for New Power Binning Sheet format.
'20181004: Modified to store the D and E value of SRAM in power binning table in performance mode, ex: SRAM_ANE.
'20180913: Modified for Multiple Power Binning.
Public Function Parsing_Row_Column(sheetName As String, idxSheet_PwrBin As Integer)
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim row As Long
    Dim col As Long
    Dim i As Integer
    Dim isSheetFound As Boolean
    Dim row_of_title As Integer
    Dim got_correct_header As Boolean
    Dim enableRowParsing As Boolean
    '''for basic columns
    Dim col_Binned_Mode As Integer
    Dim col_Other_Mode As Integer
    Dim col_Other As Integer
    Dim col_Value As Integer
    '''variables
    Dim col_Temp As Integer
    Dim idx_Temp As Integer
    Dim strTemp As String
    Dim Cell_content As String
    Dim idx_Binned_Mode As Integer
    Dim idx_Other_Mode As Integer
    Dim dict_check_BinnedMode_OtherMode As New Dictionary
    Dim str_Cell As String
    Dim split_content() As String
On Error GoTo errHandler
    '''*****************************************************************'''
    '''//Check if the sheet exists
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        PwrBin_Sheet(idxSheet_PwrBin).sheetName = sheetName
        enableRowParsing = True
        got_correct_header = False
        dict_check_BinnedMode_OtherMode.RemoveAll
        col_Other_Mode = 0
        
        '''//Check row of the header in the table.
        '''Get the columns for the diverse coefficient.
        For row = 1 To MaxRow
            For col = 1 To maxcol
                If UCase(ws_def.Cells(row, col)) = "BINNED_MODE" Then
                    col_Binned_Mode = col
                    row_of_title = row
                End If
                
                '''//Check if format of the header is correct
                If row_of_title > 0 Then
                    If UCase(ws_def.Cells(row_of_title, col)) = "OTHER_MODE" Then
                        col_Other_Mode = col
                    ElseIf UCase(ws_def.Cells(row_of_title, col)) = "OTHER" Then
                        col_Other = col
                    ElseIf UCase(ws_def.Cells(row_of_title, col)) = "VALUE" Then
                        col_Value = col
                    End If
                End If
            Next col
            
            If col_Binned_Mode > 0 Then
                If col_Other > 0 And col_Value > 0 Then
                    got_correct_header = True
                    Exit For
                Else
                    got_correct_header = False

                    If col_Other = 0 Then
                        TheExec.Datalog.WriteComment "Column OTHER doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column OTHER doesn't exist in header of " & sheetName & ". Error!!!"
                    End If

                    If col_Value = 0 Then
                        TheExec.Datalog.WriteComment "Column VALUE doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column VALUE doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    Exit For
                End If
            End If
        Next row
        
        '''****************************************************************************************'''
        '''//Coefficient of CorePower: from column "Binned_Mode" to column before "Other_Mode".
        '''//Coefficient of OtherRail: from column "Other_Mode" to column before "Other".
        '''****************************************************************************************'''
        If got_correct_header = True Then
            For col = 1 To maxcol
                '''init
                strTemp = ""
                str_Cell = UCase(ws_def.Cells(row_of_title, col).Value)
                
                '''20201211: Modified to parse new format "Binned_Mode","Bin Voltage (mV)","IDS (mA)","Vdd0 (mV)","Vdd1 (mV)" of the header, requested by C651 Toby and Ben.
                If str_Cell <> "" Then
                    split_content = Split(UCase(ws_def.Cells(row_of_title, col).Value), "(")
                    
                    '''//Get header of the column.
                    strTemp = UCase(Trim(split_content(0)))
                    
                    '''20201211: Modified to parse the powerbinning table without Other_Mode, requested by C651 Toby and Ben.
                    If col >= col_Binned_Mode And col < col_Other And strTemp <> "" Then
                        If col_Other_Mode > 0 And col >= col_Other_Mode Then  '''Other_Mode
                            If dict_Other_Mode_Ratio2Column.Exists(strTemp) = True Then
                                If col = dict_Other_Mode_Ratio2Column.Item(strTemp) Then
                                    enableRowParsing = enableRowParsing And True
                                Else
                                    enableRowParsing = enableRowParsing And False
                                    TheExec.Datalog.WriteComment "row" & row & ", col" & col & " has duplicate name " & strTemp & " in header of " & sheetName & ". Error!!!"
                                    TheExec.ErrorLogMessage "row" & row & ", col" & col & " has duplicate name " & strTemp & " in header of " & sheetName & ". Error!!!"
                                End If
                            Else
                                enableRowParsing = enableRowParsing And True
                                ReDim Preserve Other_Ratio_Name(dict_Other_Mode_Ratio2Idx.Count) As String
                                Other_Ratio_Name(dict_Other_Mode_Ratio2Idx.Count) = strTemp
                                
                                dict_Other_Mode_Column2Ratio.Add col, strTemp
                                dict_Other_Mode_Ratio2Column.Add strTemp, col
                                dict_Other_Mode_Ratio2Idx.Add strTemp, dict_Other_Mode_Ratio2Idx.Count
                            End If
                        Else '''Binned_Mode
                            If dict_Binned_Mode_Ratio2Column.Exists(strTemp) = True Then
                                If col = dict_Binned_Mode_Ratio2Column.Item(strTemp) Then
                                    enableRowParsing = enableRowParsing And True
                                Else
                                    enableRowParsing = enableRowParsing And False
                                    TheExec.Datalog.WriteComment "row" & row & ", col" & col & " has duplicate name " & strTemp & " in header of " & sheetName & ". Error!!!"
                                    TheExec.ErrorLogMessage "row" & row & ", col" & col & " has duplicate name " & strTemp & " in header of " & sheetName & ". Error!!!"
                                End If
                            Else
                                enableRowParsing = enableRowParsing And True
                                ReDim Preserve Binned_Ratio_Name(dict_Binned_Mode_Ratio2Idx.Count) As String
                                Binned_Ratio_Name(dict_Binned_Mode_Ratio2Idx.Count) = strTemp
                                
                                dict_Binned_Mode_Column2Ratio.Add col, strTemp                              '''column -> Ratio(Name).
                                dict_Binned_Mode_Ratio2Column.Add strTemp, col                              '''Ratio(Name) -> column.
                                dict_Binned_Mode_Ratio2Idx.Add strTemp, dict_Binned_Mode_Ratio2Idx.Count    '''Ratio(Name) -> Index of position in the array.
                            End If
                        End If
                    End If
                End If
            Next col
            
            PwrBin_Sheet(idxSheet_PwrBin).cnt_Binned_Mode = 0
            PwrBin_Sheet(idxSheet_PwrBin).cnt_Other_Mode = 0
        End If
                    
        '''//Start parsing the cells of Binned_Mode and Other Mode for each row.
        If enableRowParsing = True Then
            For row = row_of_title + 1 To MaxRow
                '''****************************************************************************************'''
                '''//Parsing coefficient values for the CorePower in the column "Binned_Mode".
                '''****************************************************************************************'''
                If dict_Binned_Mode_Ratio2Idx.Count > 0 Then
                    If ws_def.Cells(row, col_Binned_Mode) <> "" And (Not ws_def.Cells(row, col_Binned_Mode) Like "formula*") Then
                        '''//Check if any duplicate Binned_Mode or Other_Mode exists, requested by PCLINZG and SYLIUU.
                        If dict_check_BinnedMode_OtherMode.Exists(UCase(ws_def.Cells(row, col_Binned_Mode))) = False Then
                            dict_check_BinnedMode_OtherMode.Add UCase(ws_def.Cells(row, col_Binned_Mode)), row
                        Else
                            TheExec.Datalog.WriteComment "row" & row & ", col" & col & " has duplicate Binned_Mode " & Cell_content & " in sheet " & sheetName & ". Error!!!"
                            TheExec.ErrorLogMessage "row" & row & ", col" & col & " has duplicate Binned_Mode " & Cell_content & " in sheet " & sheetName & ". Error!!!"
                        End If
                                                
                        '''//Get index of Binned_Mode from count (index starts from 0).
                        idx_Binned_Mode = PwrBin_Sheet(idxSheet_PwrBin).cnt_Binned_Mode
                        
                        '''//Redim array size to store Binned_Mode.
                        ReDim Preserve PwrBin_Sheet(idxSheet_PwrBin).Binned_Mode(idx_Binned_Mode)
                        ReDim Preserve PwrBin_Sheet(idxSheet_PwrBin).Binned_Mode(idx_Binned_Mode).Ratio(dict_Binned_Mode_Ratio2Idx.Count - 1) As Variant

                        '''//Store all factors for Binned_Mode.
                        For i = 0 To UBound(Binned_Ratio_Name)
                            '''//Get column of the selected Ratio (factor).
                            col_Temp = dict_Binned_Mode_Ratio2Column.Item(Binned_Ratio_Name(i))
                            idx_Temp = dict_Binned_Mode_Ratio2Idx.Item(Binned_Ratio_Name(i))
                        
                            '''//Get content in the cell.
                            Cell_content = UCase(Trim(ws_def.Cells(row, col_Temp)))
                            
                            If Binned_Ratio_Name(i) = UCase("Binned_Mode") Then
                                PwrBin_Sheet(idxSheet_PwrBin).Binned_Mode(idx_Binned_Mode).Pmode = Cell_content
                                PwrBin_Sheet(idxSheet_PwrBin).Binned_Mode(idx_Binned_Mode).Ratio(idx_Temp) = Cell_content
                            ElseIf Binned_Ratio_Name(i) Like UCase("Bin Voltage") Or Binned_Ratio_Name(i) Like UCase("IDS") Then
                                '''//For temporary Binned Voltage and IDS info from columns: "Bin Voltage" and "IDS".
                                PwrBin_Sheet(idxSheet_PwrBin).Binned_Mode(idx_Binned_Mode).Ratio(idx_Temp) = CStr(Cell_content)
                            Else
                                PwrBin_Sheet(idxSheet_PwrBin).Binned_Mode(idx_Binned_Mode).Ratio(idx_Temp) = CDbl(Cell_content)
                            End If
                        Next i
                        
                        '''//Updated count of Binned_Mode. (count from 1)
                        PwrBin_Sheet(idxSheet_PwrBin).cnt_Binned_Mode = PwrBin_Sheet(idxSheet_PwrBin).cnt_Binned_Mode + 1
                        
                        '''//Check "Bin Voltage" of "Binned_Mode".
                        Call check_cell_Bin_Voltage_for_parsing_Power_Binning_table(sheetName, ws_def.Cells(row, col_Binned_Mode), CStr(PwrBin_Sheet(idxSheet_PwrBin).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("Bin Voltage"))))
                        
                        '''//Check "IDS"
                        Call check_cell_Bin_Voltage_for_parsing_Power_Binning_table(sheetName, ws_def.Cells(row, col_Binned_Mode), CStr(PwrBin_Sheet(idxSheet_PwrBin).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("IDS"))))
                    End If
                End If
                
                '''****************************************************************************************'''
                '''//Parsing coefficient values the OtherRail in the column "Other_Mode".
                '''****************************************************************************************'''
                If dict_Other_Mode_Ratio2Idx.Count > 0 Then
                    If ws_def.Cells(row, col_Other_Mode) <> "" And (Not ws_def.Cells(row, col_Other_Mode) Like "formula*") Then
                        '''//Check if any duplicate Binned_Mode or Other_Mode exists, requested by PCLINZG and SYLIUU.
                        If dict_check_BinnedMode_OtherMode.Exists(UCase(ws_def.Cells(row, col_Other_Mode))) = False Then
                            dict_check_BinnedMode_OtherMode.Add UCase(ws_def.Cells(row, col_Other_Mode)), row
                        Else
                            TheExec.Datalog.WriteComment "row" & row & ", col" & col & " has duplicate Other_Mode " & Cell_content & " in sheet " & sheetName & ". Error!!!"
                            TheExec.ErrorLogMessage "row" & row & ", col" & col & " has duplicate Other_Mode " & Cell_content & " in sheet " & sheetName & ". Error!!!"
                        End If
                                                
                        '''//Get index of Other_Mode from count (index start from 0).
                        idx_Other_Mode = PwrBin_Sheet(idxSheet_PwrBin).cnt_Other_Mode
                        
                        '''//Redim array size to store Binned_Mode.
                        ReDim Preserve PwrBin_Sheet(idxSheet_PwrBin).Other_Mode(idx_Other_Mode)
                        ReDim Preserve PwrBin_Sheet(idxSheet_PwrBin).Other_Mode(idx_Other_Mode).Ratio(dict_Other_Mode_Ratio2Idx.Count - 1) As Variant

                        '''//Store all factors for Binned_Mode.
                        For i = 0 To UBound(Other_Ratio_Name)
                            '''//Get column of the selected Ratio (factor).
                            col_Temp = dict_Other_Mode_Ratio2Column.Item(Other_Ratio_Name(i))
                            idx_Temp = dict_Other_Mode_Ratio2Idx.Item(Other_Ratio_Name(i))
                        
                            '''//Get content in the cell.
                            Cell_content = UCase(Trim(ws_def.Cells(row, col_Temp)))
                            
                            If Other_Ratio_Name(i) = UCase("Other_Mode") Then
                                PwrBin_Sheet(idxSheet_PwrBin).Other_Mode(idx_Other_Mode).Pmode = Cell_content
                                PwrBin_Sheet(idxSheet_PwrBin).Other_Mode(idx_Other_Mode).Ratio(idx_Temp) = Cell_content
                            ElseIf Other_Ratio_Name(i) Like UCase("Bin Voltage") Or Binned_Ratio_Name(i) Like UCase("IDS") Then
                                '''//For temporary Binned Voltage and IDS info from columns: "Bin Voltage" and "IDS".
                                PwrBin_Sheet(idxSheet_PwrBin).Other_Mode(idx_Other_Mode).Ratio(idx_Temp) = Cell_content
                            Else
                                PwrBin_Sheet(idxSheet_PwrBin).Other_Mode(idx_Other_Mode).Ratio(idx_Temp) = CDbl(Cell_content)
                            End If
                        Next i
                        
                        '''//Updated count of Other Mode.
                        PwrBin_Sheet(idxSheet_PwrBin).cnt_Other_Mode = PwrBin_Sheet(idxSheet_PwrBin).cnt_Other_Mode + 1
                        
                        '''//Check "Bin Voltage" of "Binned_Mode".
                        Call check_cell_Bin_Voltage_for_parsing_Power_Binning_table(sheetName, ws_def.Cells(row, col_Other_Mode), CStr(PwrBin_Sheet(idxSheet_PwrBin).Binned_Mode(idx_Other_Mode).Ratio(PwB_Other_Mode_Ratio2Idx("Bin Voltage"))))

                        '''//Check "IDS"
                        Call check_cell_IDS_for_parsing_Power_Binning_table(sheetName, ws_def.Cells(row, col_Other_Mode), CStr(PwrBin_Sheet(idxSheet_PwrBin).Binned_Mode(idx_Other_Mode).Ratio(PwB_Other_Mode_Ratio2Idx("IDS"))))
                    End If
                End If

                '''//Parsing the offset and spec values for each sheet of power binning
                If IsEmpty(ws_def.Cells(row, col_Other)) = False Then
                    Cell_content = Trim((ws_def.Cells(row, col_Other)))

                    If UCase(ws_def.Cells(row, col_Other)) Like "*OFFSET*" Then
                        PwrBin_Sheet(idxSheet_PwrBin).Offset = CDbl(ws_def.Cells(row, col_Value))
                    ElseIf UCase(ws_def.Cells(row, col_Other)) Like "*SPEC*" Then
                        PwrBin_Sheet(idxSheet_PwrBin).spec = CDbl(ws_def.Cells(row, col_Value))
                    End If
                End If ''' If IsEmpty(ws_def.Cells(i, col_Other)) = False Then
            Next row '''Move to next row
        Else
            TheExec.Datalog.WriteComment sheetName & " doesn't have the correct header for Parsing_Row_Column. Error!!!"
            TheExec.ErrorLogMessage sheetName & " doesn't have the correct header for Parsing_Row_Column. Error!!!"
        End If '''If enableRowParsing = True
    End If '''If isSheetFound = True
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Parsing_Row_Column"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Parsing_Row_Column"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201110: Modified to revised the vbt code for parsing PowerBinning sheets with the new format.
Public Function PwB_Binned_Mode_Ratio2Idx(InputRatioName As String) As Long
On Error GoTo errHandler
    If dict_Binned_Mode_Ratio2Idx.Exists(UCase(InputRatioName)) = True Then
        PwB_Binned_Mode_Ratio2Idx = dict_Binned_Mode_Ratio2Idx.Item(UCase(InputRatioName))
    Else
        PwB_Binned_Mode_Ratio2Idx = -1
        TheExec.Datalog.WriteComment InputRatioName & " doesn't exist in dict_Binned_Mode_Ratio2Idx, It can't get PwB_Binned_Mode_Ratio2Idx. Error!!!"
        TheExec.ErrorLogMessage InputRatioName & " doesn't exist in dict_Binned_Mode_Ratio2Idx, It can't get PwB_Binned_Mode_Ratio2Idx. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of PwB_Binned_Ratio2Idx"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of PwB_Binned_Ratio2Idx"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201110: Modified to revised the vbt code for parsing PowerBinning sheets with the new format.
Public Function PwB_Other_Mode_Ratio2Idx(InputRatioName As String) As Long
On Error GoTo errHandler
    If dict_Other_Mode_Ratio2Idx.Exists(UCase(InputRatioName)) = True Then
        PwB_Other_Mode_Ratio2Idx = dict_Other_Mode_Ratio2Idx.Item(UCase(InputRatioName))
    Else
        PwB_Other_Mode_Ratio2Idx = -1
        TheExec.Datalog.WriteComment InputRatioName & " doesn't exist in dict_Other_Mode_Ratio2Idx, It can't get PwB_Other_Mode_Ratio2Idx. Error!!!"
        TheExec.ErrorLogMessage InputRatioName & " doesn't exist in dict_Other_Mode_Ratio2Idx, It can't get PwB_Other_Mode_Ratio2Idx. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of PwB_Other_Ratio2Idx"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of PwB_Other_Ratio2Idx"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201110: Created to check cell contents of "Bin Voltage" for Binned_Mode and Other_Mode in Power_Binning table.
Public Function check_cell_Bin_Voltage_for_parsing_Power_Binning_table(sheetName As String, str_PwrBin_Mode As String, str_Bin_Voltage As String)
    Dim i As Integer
    Dim strTemp As String
    Dim strAry_powerDomain() As String
On Error GoTo errHandler
    '''//Check "Bin Voltage" of "Binned_Mode".
    If LCase(str_Bin_Voltage) Like "max(*vdd*,*vdd*)" Then
        strTemp = Replace(UCase(str_Bin_Voltage), "MAX(", "")
        strTemp = Replace(UCase(strTemp), ")", "")
        strAry_powerDomain = Split(strTemp, ",")

        If LCase(str_PwrBin_Mode) Like "sram_*" Then
            '''Do nothing
        Else
            TheExec.Datalog.WriteComment "Sheet:" & sheetName & ", Binned_Mode:" & str_PwrBin_Mode & ", has the incorrect format. It should have the prefix SRAM_XXX. Error!!!"
            TheExec.ErrorLogMessage "Sheet:" & sheetName & ", Binned_Mode:" & str_PwrBin_Mode & ", has the incorrect format. It should have the prefix SRAM_XXX. Error!!!"
        End If
    Else
        ReDim strAry_powerDomain(0)
        strAry_powerDomain(0) = str_Bin_Voltage
    End If

    For i = 0 To UBound(strAry_powerDomain)
        If VddBinStr2Enum(UCase(strAry_powerDomain(i))) <= cntVddbinPin Then
            '''Do nothing
        Else
            TheExec.Datalog.WriteComment "Sheet:" & sheetName & ", Mode:" & str_PwrBin_Mode & ", Bin Voltage: " & str_Bin_Voltage & ", is not the correct BinCut powerDomain. Error!!!"
            TheExec.ErrorLogMessage "Sheet:" & sheetName & ", Mode:" & str_PwrBin_Mode & ", Bin Voltage: " & str_Bin_Voltage & ", is not the correct BinCut powerDomain. Error!!!"
        End If
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of check_cell_Bin_Voltage_for_parsing_Power_Binning_table"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of check_cell_Bin_Voltage_for_parsing_Power_Binning_table"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201110: Created to check cell contents of "IDS" for Binned_Mode and Other_Mode in Power_Binning table.
Public Function check_cell_IDS_for_parsing_Power_Binning_table(sheetName As String, str_PwrBin_Mode As String, str_IDS As String)
    Dim i As Integer
    Dim strTemp As String
    Dim strAry_powerDomain() As String
On Error GoTo errHandler
    '''//Check "IDS"
    If LCase(str_IDS) Like "vdd*+*vdd*" Then
        strAry_powerDomain = Split(str_IDS, "+")
    Else
        ReDim strAry_powerDomain(0)
        strAry_powerDomain(0) = str_IDS
    End If

    For i = 0 To UBound(strAry_powerDomain)
        If VddBinStr2Enum(UCase(strAry_powerDomain(i))) <= cntVddbinPin Then
            '''Do nothing
        Else
            TheExec.Datalog.WriteComment "Sheet:" & sheetName & ", Binned_Mode:" & str_PwrBin_Mode & ", IDS: " & str_IDS & ", is not the correct BinCut powerDomain. Error!!!"
            TheExec.ErrorLogMessage "Sheet:" & sheetName & ", Binned_Mode:" & str_PwrBin_Mode & ", IDS: " & str_IDS & ", is not the correct BinCut powerDomain. Error!!!"
        End If
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of check_cell_IDS_for_parsing_Power_Binning_table"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of check_cell_IDS_for_parsing_Power_Binning_table"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201116: Modified to check "LCase(performance_mode) Like "sram_m*##*".
'20201112: Modified to check if powerDomain is BinCut CorePower or OtherRail.
'20201112: Modified to rename the vbt function "ParsingBinVolt" as "get_Voltage_for_PowerBinning".
'20201112: Modified to revise the vbt code for Other_Mode, ex: "MPS001", "MES001".
'20200106: Modified to remove the ErrorLogMessage.
'20191127: Modified for the revised InitVddBinTable.
'20190813: Modified to get p_mode for parsing the string of Binned_Mode in powerbinning tables.
'20190311: Modified to add voltage_Temp for GradeVdd (product voltage).
Public Function get_Voltage_for_PowerBinning(performance_mode As String, site As Variant, bin_number As Long, strPwrbinBinVoltage As Variant) As Double
    Dim strAry_powerDomain() As String
    Dim split_content() As String
    Dim voltage_Temp As Double
    Dim operator_Temp As String
    Dim i As Long
    Dim p_mode As Integer
    Dim strTemp As String
On Error GoTo errHandler
    '''//[Step1] Find the keyword to split the string from PowerBinning Ratio "Bin Voltage".
    If InStr(LCase(strPwrbinBinVoltage), "max") > 0 Then '''ex: "Max(VDD_GPU,VDD_SRAM_GPU)"
        operator_Temp = "max"
    Else
        operator_Temp = " "
    End If
    
    '''//[Step2] strAry_powerDomain stores powerDomains from from "Bin Voltage".
    Split_special_string_for_calc CStr(strPwrbinBinVoltage), operator_Temp, strAry_powerDomain
    
    '''//[Step3] Check if string from "Binned_Mode" or "Other_Mode" contains any keyword about performance_mode.
    If LCase(performance_mode) Like "m*##*" Then
        If LCase(performance_mode) Like "m*##*_*" Then '''ex: "MS001_GPU"
            split_content = Split(performance_mode, "_")
            strTemp = UCase(Trim(split_content(0)))
        Else
            If Len(Trim(performance_mode)) = 6 Then '''ex: "MPS001", "MES001".
                strTemp = UCase(Trim(Mid(performance_mode, 1, 2) & Mid(performance_mode, 4, 3)))
            Else
                strTemp = UCase(Trim(performance_mode))
            End If
        End If
    ElseIf LCase(performance_mode) Like "sram_m*##*" Then '''ex: "SRAM_MG001".
        strTemp = UCase(Trim(Replace(UCase(performance_mode), "SRAM_", "")))
    Else
        If UCase(performance_mode) Like "VDD_*" Then '''ex: "VDD_SRAM_SOC".
            strTemp = UCase(Trim(performance_mode))
        Else '''ex: "SRAM_SOC", "FIXED", "LOW".
            strTemp = UCase(Trim("VDD_" & performance_mode))
        End If
    End If
            
    '''//[Step4] Get p_mode for parsing performance_mode from column "Binned_Mode" or "Other_Mode" in Powerbinning tables.
    If VddbinPmodeDict.Exists(UCase(strTemp)) Then
        p_mode = VddBinStr2Enum(UCase(strTemp))
    Else
        strTemp = ""
        TheExec.Datalog.WriteComment "get_Voltage_for_PowerBinning doesn't get any correct performance_mode from " & performance_mode & ". Error!!!"
        TheExec.ErrorLogMessage "get_Voltage_for_PowerBinning doesn't get any correct performance_mode from " & performance_mode & ". Error!!!"
    End If
    
    '''//[Step5] Get Efuse product voltage by p_mode for each powerDomain.
    If strTemp <> "" Then
        For i = 0 To UBound(strAry_powerDomain)
            '''//Check if powerDomain is BinCut CorePower or OtherRail.
            If dict_IsCorePowerInBinCutFlowSheet.Exists(UCase(strAry_powerDomain(i))) Then
                '''//Get Efuse product voltage.
                If dict_IsCorePowerInBinCutFlowSheet.Item(UCase(strAry_powerDomain(i))) Then '''CorePower
                    voltage_Temp = VBIN_RESULT(p_mode).GRADEVDD(site) '''unit: mV
                Else '''OtherRail
                    voltage_Temp = BinCut(p_mode, bin_number).OTHER_PRODUCT_RAIL(VddBinStr2Enum(Trim(strAry_powerDomain(i)))) '''unit: mV
                End If
            Else
                voltage_Temp = 0
                TheExec.Datalog.WriteComment strPwrbinBinVoltage & " from Bin Voltage of " & performance_mode & " in PowerBinning table is not BinCut CorePower or OtherRail for get_Voltage_for_PowerBinning. Error!!!"
                TheExec.ErrorLogMessage strPwrbinBinVoltage & " from Bin Voltage of " & performance_mode & " in PowerBinning table is not BinCut CorePower or OtherRail for get_Voltage_for_PowerBinning. Error!!!"
            End If
            
            If voltage_Temp > get_Voltage_for_PowerBinning Then get_Voltage_for_PowerBinning = voltage_Temp
        Next i
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of get_Voltage_for_PowerBinning"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of get_Voltage_for_PowerBinning"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201112: Modified to check if powerDomain is BinCut CorePower or OtherRail.
'20201112: Modified to rename the vbt function "ParsingIDS" as "get_IDS_for_PowerBinning".
'20200121: Modified for VDD_FIXED_GRP and VDD_LOW_GRP.
'20200114: Modified to check if powerDomain exists in domain2pinDict or pin2domainDict.
'20190528: Modified to get IDS values from the new IDS datatype.
Public Function get_IDS_for_PowerBinning(site As Variant, str_IDS As Variant) As Double
    Dim operator_Temp As String
    Dim split_content() As String
    Dim idsTemp As Double
    Dim i As Long
    Dim powerDomain As String
On Error GoTo errHandler
    '''//init
    idsTemp = 0
    
    If InStr(str_IDS, "+") > 0 Then
        operator_Temp = "+"
    Else
        operator_Temp = " "
    End If
    
    split_content = Split(str_IDS, operator_Temp)
    
    For i = 0 To UBound(split_content)
        '''//Check if powerDomain exists in domain2pinDict or pin2domainDict.
        If domain2pinDict.Exists(UCase(split_content(i))) Then
            powerDomain = UCase(split_content(i))
        ElseIf pin2domainDict.Exists(UCase(split_content(i))) Then
            powerDomain = VddbinPin2Domain(UCase(split_content(i)))
        Else
            powerDomain = ""
            TheExec.Datalog.WriteComment split_content(i) & " is not BinCut CorePower or OtherRail for get_IDS_for_PowerBinning. Please check PowerBinning table. Error!!!"
            TheExec.ErrorLogMessage split_content(i) & " is not BinCut CorePower or OtherRail for get_IDS_for_PowerBinning. Please check PowerBinning table. Error!!!"
        End If
        
        If powerDomain <> "" Then
            '''ToDo: Maybe we have to modify this for CP1 and non-CP1...
            idsTemp = IDS_for_BinCut(VddBinStr2Enum(Trim(powerDomain))).Real(site) '''unit: mA
            
            '''//If idsTemp > ParsingIDS, get_IDS_for_PowerBinning = idsTemp
            If operator_Temp = "+" Then
                get_IDS_for_PowerBinning = get_IDS_for_PowerBinning + idsTemp
            Else
                get_IDS_for_PowerBinning = idsTemp
            End If
        Else
            get_IDS_for_PowerBinning = 0
        End If
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of get_IDS_for_PowerBinning"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of get_IDS_for_PowerBinning"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210715: Modified to check if voltage_Binned_Mode<=0 for calculate_power_for_binned_mode.
'20210504: Modified to get ratios for powerBinning.
'20210503: Modified to revise the vbt code for the new formulas from C651 Si and Toby.
'20201215: C651 Toby and Ben updated the formulas for Binned_Mode.
'20201125: Modified to revise the branches for different formulas, requested by PCLINZG.
'20201123: Modified to revise the vbt code for Exp2, Exp1, and Exp0 for PowerBinning formulas.
'20201112: Created to calculate power for Binned_Mode of PowerBinning.
'20201029: C651 Toby proposed the new format of PowerBinning table.
Public Function calculate_power_for_binned_mode(idx_Sheet As Integer, idx_Binned_Mode As Integer, voltage_Binned_Mode As Double, ids_Binned_Mode As Double) As Double
    Dim ratio_A As Double
    Dim ratio_B As Double
    Dim ratio_C As Double
    Dim ratio_J As Double
    Dim ratio_K As Double
    Dim ratio_Base As Double
    Dim ratio_Vdd1 As Double
    Dim ratio_Vdd0 As Double
    Dim ratio_Exp2 As Double
    Dim ratio_Exp1 As Double
    Dim ratio_Exp0 As Double
On Error GoTo errHandler
    '''//Check if Efuse product voltage of binned_mode<=0.
    '''20210715: Modified to check if voltage_Binned_Mode<=0 for calculate_power_for_binned_mode.
    If voltage_Binned_Mode <= 0 Then
        TheExec.Datalog.WriteComment "Warning!!! Sheet:" & PwrBin_Sheet(idx_Sheet).sheetName & ",Binned_Mode:" & PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Pmode & ", product voltage <= 0 might cause ErrorHandler issues in calculate_power_for_binned_mode. Please check test results of the binned_mode in BV test instance. Warning!!!"
    End If

    '''//Get ratio C
    ratio_C = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("C"))

    If dict_Binned_Mode_Ratio2Idx.Exists(UCase("A")) = True And dict_Binned_Mode_Ratio2Idx.Exists(UCase("B")) = True And dict_Binned_Mode_Ratio2Idx.Exists(UCase("C")) = True Then
        '''=========================================================================================='''
        '''//Formula: Pbinned_mode=(A+B*IDS)*((Vbin-C)^2)/(10^6)
        '''=========================================================================================='''
        ratio_A = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("A"))
        ratio_B = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("B"))
        calculate_power_for_binned_mode = (ratio_A + ratio_B * ids_Binned_Mode) * ((voltage_Binned_Mode - ratio_C) ^ 2) / 1000000
                                            
    ElseIf dict_Binned_Mode_Ratio2Idx.Exists(UCase("Vdd0")) = True And dict_Binned_Mode_Ratio2Idx.Exists(UCase("Vdd1")) = True Then
        '''=========================================================================================='''
        '''//20201215: C651 Toby and Ben updated the formulas for Binned_Mode.
        '''//PV is Efuse product voltage of Binned_Mode.
        '''PV-C <=Vdd1       ->  Formula: Pbinned_mode=J*((Vbin-C)^Exp2)/10^6 + K*IDS*((Vbin-C)/Vdd1)^Exp1)
        '''Vdd1<PV-C<=Vdd0  ->  Formula: Pbinned_mode=J*((Vbin-C)^Exp2)/10^6 + K*IDS*((Vbin-C)/Vdd1)^Exp0)
        '''PV-C>Vdd0        ->  Formula: Pbinned_mode=J*((Vbin-C)^Exp2)/10^6 + K*IDS*((Vbin-C)*Vdd0)/(Vdd1^Exp0))*(Base^(Vbin-C-Vdd0))
        '''=========================================================================================='''
        ratio_J = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("J"))
        ratio_K = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("K"))
        ratio_Base = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("Base"))
        ratio_Vdd1 = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("Vdd1"))
        ratio_Vdd0 = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("Vdd0"))
        ratio_Exp2 = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("Exp2"))
        ratio_Exp1 = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("Exp1"))
        ratio_Exp0 = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).Ratio(PwB_Binned_Mode_Ratio2Idx("Exp0"))
        
        If CDec(voltage_Binned_Mode - ratio_C) <= CDec(ratio_Vdd1) Then '''PV<=Vdd1
            calculate_power_for_binned_mode = (ratio_J / 1000000) * (voltage_Binned_Mode - ratio_C) ^ ratio_Exp2 + _
                                ratio_K * ids_Binned_Mode * (((voltage_Binned_Mode - ratio_C) / ratio_Vdd1) ^ ratio_Exp1)
                                                
        ElseIf CDec(voltage_Binned_Mode - ratio_C) > CDec(ratio_Vdd1) And CDec(voltage_Binned_Mode - ratio_C) <= CDec(ratio_Vdd0) Then '''Vdd1<PV<=Vdd0
            calculate_power_for_binned_mode = (ratio_J / 1000000) * (voltage_Binned_Mode - ratio_C) ^ ratio_Exp2 + _
                                ratio_K * ids_Binned_Mode * (((voltage_Binned_Mode - ratio_C) / ratio_Vdd1) ^ ratio_Exp0)
                                                
        ElseIf CDec(voltage_Binned_Mode - ratio_C) > CDec(ratio_Vdd0) Then '''PV>Vdd0
            calculate_power_for_binned_mode = (ratio_J / 1000000) * (voltage_Binned_Mode - ratio_C) ^ ratio_Exp2 + _
                                ratio_K * ids_Binned_Mode * ((voltage_Binned_Mode - ratio_C) * ratio_Vdd0 / (ratio_Vdd1 ^ ratio_Exp0)) * (ratio_Base ^ (voltage_Binned_Mode - ratio_C - ratio_Vdd0))
        End If
    Else
        calculate_power_for_binned_mode = 0
        TheExec.Datalog.WriteComment "Sheet:" & PwrBin_Sheet(idx_Sheet).sheetName & ", it doesn't have the Ratio A/B/C or Vdd0/Vdd1 in the header. Please check PowerBinning tables and the formula calculate_power_for_binned_mode. Error!!!"
        TheExec.ErrorLogMessage "Sheet:" & PwrBin_Sheet(idx_Sheet).sheetName & ", it doesn't have the Ratio A/B/C or Vdd0/Vdd1 in the header. Please check PowerBinning tables and the formula calculate_power_for_binned_mode. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of calculate_power_for_binned_mode"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of calculate_power_for_binned_mode"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201211: C651 Toby and Ben updated the formulas for Other_Mode.
'20201112: Created to calculate power for Other_Mode of PowerBinning.
'20201029: C651 Toby proposed the new format of PowerBinning table.
Public Function calculate_power_for_other_mode(idx_Sheet As Integer, idx_Other_Mode As Integer, voltage_Other_Mode As Double, ids_Other_Mode As Double) As Double
On Error GoTo errHandler
    If dict_Other_Mode_Ratio2Idx.Exists(UCase("D")) = True And dict_Other_Mode_Ratio2Idx.Exists(UCase("E")) = True Then
        '''=========================================================================================='''
        '''//Formula: Pother_mode=D+E*IDS
        '''=========================================================================================='''
        calculate_power_for_other_mode = PwrBin_Sheet(idx_Sheet).Other_Mode(idx_Other_Mode).Ratio(PwB_Other_Mode_Ratio2Idx("D")) + ids_Other_Mode * PwrBin_Sheet(idx_Sheet).Other_Mode(idx_Other_Mode).Ratio(PwB_Other_Mode_Ratio2Idx("E"))
    Else
        calculate_power_for_other_mode = 0
        TheExec.Datalog.WriteComment PwrBin_Sheet(idx_Sheet).sheetName & " doesn't have the correct Ratio for Other_Mode. Please check PowerBinning tables and the formula for calculate_power_for_other_mode. Error!!!"
        TheExec.ErrorLogMessage PwrBin_Sheet(idx_Sheet).sheetName & " doesn't have the correct Ratio for Other_Mode. Please check PowerBinning tables and the formula for calculate_power_for_other_mode. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of calculate_power_for_other_mode"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of calculate_power_for_other_mode"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210908: Modified to merge the vbt code of the vbt function Apply_hvcctestcondition_InFlowSheet into the vbt function Apply_testcondition_InFlowSheet, as discussed with TSMC ZYLINI.
'20210906: Modified to remove the redundant argument "ByRef voltage_SelsrmBitCalc As SiteDouble".
'20210813: Modified to revise the vbt code for postBinCut or outsideBinCut instance names with keyword "*_binresult_" for the vbt initialize_inst_info and Apply_testcondition_InFlowSheet.
'20210812: As per discussion with TSMC ZYLINI, he suggested us to replace the keyword of testCondition "*Evaluate*Bin*" with "*Bin*Result*" if p_mode is tested.
'20210802: Modified to merge the vbt functions Calculate_LVCC_for_CorePower and Calculate_HVCC_for_CorePower into the vbt function Calculate_Voltage_for_CorePower.
'20210729: Modified to remove the redundant argument "ids_current As SiteDouble" from the vbt functions Apply_testcondition_InFlowSheet and Calculate_LVCC_for_CorePower.
'20210505: Modified to print powerDomain for the error message.
'20210302: Modified to optimize the keyword replacement to "M*### Bin Result".
'20210126: Modified to revise the vbt code for DevChar.
'20201211: Modified to decide the replacement of keyword in the testCondition by the flag "is_BinCutJob_for_StepSearch".
'20201119: Modified to check "bincutJobName = "cp1" to replace keyword of testCondition with "Bin Result" for postBinCut and outsideBinCut, requested by PCLINZG.
'20201116: Modified to use "testCondition = LCase(Trim(Get_BinCut_TestCondition()))".
'20201029: Modified to replace the argument "corepower as Integer" with "idxPin_CorePower As Integer".
'20201029: Modified to use "Public Type Instance_Info".
'20200806: Modified to use siteDouble for Calculate_LVCC_for_CorePower, Calculate_HVCC_for_CorePower, and Calculate_Voltage_for_OtherRail.
'20200803: Modified to use the argument "ByRef voltage_SelsrmBitCalc As SiteDouble".
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20200711: Modified to change the argumens types and store BinCut payload voltages into the siteDouble array.
'20200710: Modified to merge the redundant site-loop for voltage calculation.
'20200430: Modified to print siteNumber for the site with the incorrect voltage.
'20200429: Modified to use "Get_BinCut_TestCondition" to get BinCut testCondition.
'20200106: Modified to remove the ErrorLogMessage.
'20191227: Modified to check the string of TestCondition in Non_Binning_Rail.
'20191216: Modified for "900mV (MS003)".
'20191202: Modified for the revised initVddBinCondition.
'20191127: Modified for the revised InitVddBinTable.
'20190507: Modified to add "Cdec" to avoid double format accuracy issues.
'20190424: Modified to replace keyword of the testCondition with "bin result".
'20190314: Modified for calling the function Calculate_LVCC_for_CorePower to calculate CorePower voltage.
'20190304: Modified for calling the new funciton Calculate_Extra_Voltage_for_PowerRail.
'20181009: Modified to replace keyword of BinCut testCondition with "bin result" for postBinCut and retention.
'20180921: Modified to re-assembly the non_binning_rail string with the keyword "bin result".
'20180913: Modified to prevent output values from DCVS out of range.
Public Function Apply_testcondition_InFlowSheet(inst_info As Instance_Info, idxPin_CorePower As Integer, PassBinNum As SiteLong, ByRef voltage_CorePower As SiteDouble)
    Dim site As Variant
    Dim str_testCondition As String
    Dim split_content() As String
    Dim i As Long
    Dim other_p_mode As Integer
    Dim idxPin_OtherRail As Integer
    Dim strAry_TestCondition() As String
    Dim voltageArray_Calculated() As New SiteDouble
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''<Keyword replacement of BinCut test condition>
'''20180926: Currently C651 Toby didn't define "bin result" or "product-*gb" for postBinCut and retention tests.
'''So that we define the voltage as "VBIN_RESULT(P_mode).Grade". If we get the definition from Toby, we will update this.
'''<Warning!!!>
'''ToDo: Please discuss this with C651 project DRIs to see if we can use the keyword in the instance names to decide the keyword replacement of BinCut test condition.
'''20210812: As per discussion with TSMC ZYLINI, he suggested us to replace the keyword of testCondition "*Evaluate*Bin*" with "*Bin*Result*" if p_mode is tested.
'''//==================================================================================================================================================================================//'''
    For Each site In TheExec.sites
        '''//Get BinCut testCondition for the selected powerDomain(CorePower).
        str_testCondition = LCase(Trim(Get_BinCut_TestCondition(inst_info, idxPin_CorePower, PassBinNum(site))))
        
        '''20210813: Modified to revise the vbt code for postBinCut or outsideBinCut instance names with keyword "*_binresult_" for the vbt initialize_inst_info and Apply_testcondition_InFlowSheet.
        '''*************************************************************************************'''
        '''//Keyword replacement of BinCut test condition of p_mode.
        '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
        '''Since PostBinCut_Voltage_Set_VT support BV and HBV tests, it should replace keyword of testCondition from BV with "bin result".
        '''//For the special case, ex: "900mV (MS003)", do not replace the keyword of testCondtion with with "*Bin*Result*".
        '''20181009: As the request from KTCHAN, he defined that postbincut instances must have the keyword "_binresult_".
        '''20210126: Modified to revise the vbt code for DevChar.
        '''20210302: Modified to optimize the keyword replacement to "M*### Bin Result".
        If is_BinCutJob_for_StepSearch = True Then '''only for postBinCut or outsideBincut in CP1.
            If (LCase(inst_info.inst_name) Like "*_binresult_*" And Not (str_testCondition) Like "*bin*result*" And Not (str_testCondition) Like "*#mv*") Or inst_info.is_DevChar_Running = True Then
                split_content = Split(str_testCondition, " ")
                
                '''//Check if any correct keyword of performance_mode exists...
                If VddbinPmodeDict.Exists(UCase(Trim(split_content(0)))) Then
                    '''//p_mode of non_binning CorePower.
                    If VBIN_RESULT(VddBinStr2Enum(UCase(split_content(0)))).tested = True Then
                        str_testCondition = split_content(0) & " " & "bin result"
                    End If
                End If
            End If
        End If
        '''*************************************************************************************'''
        
        '''//Use the function Calculate_LVCC_for_CorePower to calculate CorePower voltage.
        '''//Check if testCondition contains the special keyword.
        '''20210908: Modified to merge the vbt code of the vbt function Apply_hvcctestcondition_InFlowSheet into the vbt function Apply_testcondition_InFlowSheet.
        If str_testCondition Like "max*(*)" Then '''ex: "MAX(MD003 Product +10%, VDD_SRAM_SOC CP HVCC)"
            Split_special_string_for_calc str_testCondition, "max", strAry_TestCondition
            ReDim voltageArray_Calculated(UBound(strAry_TestCondition)) As New SiteDouble
            
            For i = 0 To UBound(strAry_TestCondition)
                voltageArray_Calculated(i) = 0
            Next i
            
            For i = 0 To UBound(strAry_TestCondition)
                If strAry_TestCondition(i) Like "*sram*" Then '''Calculate voltage for OtherRail, ex: "VDD_SRAM_SOC CP HVCC".
                    split_content = Split(Trim(strAry_TestCondition(i)), " ")
                    
                    If Trim(split_content(0)) Like "*sram*" Then
                        idxPin_OtherRail = VddBinStr2Enum(split_content(0))
                        
                        '''//Calculate the voltage for OtherRail.
                        Call Calculate_Voltage_for_OtherRail(inst_info, idxPin_OtherRail, PassBinNum, Trim(strAry_TestCondition(UBound(strAry_TestCondition))), voltageArray_Calculated(i))
                    Else
                        TheExec.Datalog.WriteComment strAry_TestCondition(i) & " from testCondition:" & str_testCondition & " for powerDomain:" & VddBinName(idxPin_CorePower) & " from sheet Non_Binning_Rail has the incorrect content. Error!!!"
                        'TheExec.ErrorLogMessage strAry_TestCondition(i) & " from testCondition:" & str_testCondition & " for powerDomain:" & VddBinName(idxPin_CorePower) & " from sheet Non_Binning_Rail has the incorrect content. Error!!!"
                    End If
                Else '''//Calculate voltage for CorePower.
                    split_content = Split(Trim(strAry_TestCondition(i)), " ")
                    
                    If Trim(split_content(0)) Like "m*##*" Then '''ex: "MC60A"
                        other_p_mode = VddBinStr2Enum(split_content(0))
                    Else
                        TheExec.Datalog.WriteComment "HVCC testCondition:" & str_testCondition & " for powerDomain:" & VddBinName(idxPin_OtherRail) & " from sheet Non_Binning_Rail has the incorrect content. Error!!!"
                        'TheExec.ErrorLogMessage "HVCC testCondition:" & str_testCondition & " for powerDomain:" & VddBinName(idxPin_OtherRail) & " from sheet Non_Binning_Rail has the incorrect content. Error!!!"
                    End If
                    
                    '''//Calculate the voltage for CorePower.
                    '''20210802: Modified to merge the vbt functions Calculate_LVCC_for_CorePower and Calculate_HVCC_for_CorePower into the vbt function Calculate_Voltage_for_CorePower.
                    Call Calculate_Voltage_for_CorePower(inst_info, idxPin_CorePower, PassBinNum, strAry_TestCondition(i), voltageArray_Calculated(i))
                End If
            Next i
            
            '''//Select the highest voltage for the temp_voltage().
            voltage_CorePower = WorksheetFunction.max(voltageArray_Calculated)
            
            '''//for SELSRM DSSC bit calculation
            inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(site) = voltage_CorePower(site)
        Else '''ex: "MD003 Product +10%"
            '''20210802: Modified to merge the vbt functions Calculate_LVCC_for_CorePower and Calculate_HVCC_for_CorePower into the vbt function Calculate_Voltage_for_CorePower.
            Call Calculate_Voltage_for_CorePower(inst_info, idxPin_CorePower, PassBinNum, str_testCondition, voltage_CorePower)
        End If
        
        '''//Floor/Ceiling BinCut payload voltage for BV/HBV test instances.
        If CDec(voltage_CorePower) > 0 Then '''prevent DCVS from out of range error.
            '''20210908: Modified to merge the vbt code of the vbt function Apply_hvcctestcondition_InFlowSheet into the vbt function Apply_testcondition_InFlowSheet, as discussed with TSMC ZYLINI.
            If LCase(inst_info.inst_name) Like "*_hbv" Then '''HBV
                '''//HBV instances ceiling the value of BinCut payload voltage.
                voltage_CorePower(site) = Ceiling(voltage_CorePower(site)) '''unit: mV
            ElseIf LCase(inst_info.inst_name) Like "*_bv" Then '''BV
                '''//BV instances floor the value of BinCut payload voltage.
                voltage_CorePower(site) = Floor(voltage_CorePower(site)) '''unit: mV
            Else
                TheExec.Datalog.WriteComment "instance:" & inst_info.inst_name & ", it doesn't contain any keyword about _bv or _hbv for Apply_testcondition_InFlowSheet. Error!!!"
                TheExec.ErrorLogMessage "instance:" & inst_info.inst_name & ", it doesn't contain any keyword about _bv or _hbv for Apply_testcondition_InFlowSheet. Error!!!"
            End If
        Else
            voltage_CorePower(site) = 0
            TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ",powerDomain:" & VddBinName(idxPin_CorePower) & ", BinCut payload voltage should be larger than 0. Error occurs in Apply_testcondition_InFlowSheet while " & inst_info.inst_name & ". Error!!!"
            'TheExec.ErrorLogMessage "site:" & site & ","  & VddBinName(inst_info.p_mode) & ",powerDomain:" & VddBinName(idxPin_CorePower) &  ", BinCut payload voltage should be larger than 0. Error occurs in Apply_testcondition_InFlowSheet while " & inst_info.inst_name & ". Error!!!"
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Apply_testcondition_InFlowSheet"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Apply_testcondition_InFlowSheet"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210908: Modified to add the argument "Optional enable_DynamicOffset As Boolean = False" to calculate BinCut payload voltage of the binning PowerDomain with DynamicOffset.
'20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
'20210908: Modified to remove the redundant argument "ids_current As SiteDouble" from the vbt functions Apply_hvcctestcondition_InFlowSheet.
'20210908: Modified to merge the vbt code of the vbt function HVCC_OtherRail_Pwr_Setting into the vbt function OtherRail_Pwr_Setting, as discussed with TSMC ZYLINI.
'20210824: Modified to call the vbt function Calculate_Binning_CorePower_with_DynamicOffset.
'20210824: Modified to add the argument "PassBinNum As SiteLong" to the vbt function Calculate_Binning_CorePower_with_DynamicOffset.
'20210824: Modified to remove the vbt code of checking the powerDomain of the binning performance mode.
'20210729: Modified to remove the redundant argument "ids_current As SiteDouble" from the vbt functions Apply_testcondition_InFlowSheet and Calculate_LVCC_for_CorePower.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210207: Modified to add the argument "idxPin_CorePower As Integer" to the vbt function Calculate_LVCC_for_CorePower and Calculate_HVCC_for_CorePower.
'20210126: Modified to revise the vbt code for DevChar.
'20201209: Modified to remove the argument "ByRef voltage_SelsrmBitCalc() As SiteDouble" and use "inst_info.voltage_SelsrmBitCalc" for HVCC_Set_VT.
'20201209: Modified to remove the argument "ByRef voltage_SelsrmBitCalc() As SiteDouble" and use "inst_info.voltage_SelsrmBitCalc" for Non_Binning_Pwr_Setting_VT.
'20201029: Modified to use "Public Type Instance_Info".
'20200803: Modified to use the argument "ByRef voltage_SelsrmBitCalc() As SiteDouble".
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20200713: Modified to remove the argument "IndexLevelPerSite As SiteLong" by using the function Get_PassBinNum_by_Step.
'20200711: Modified to change the argumens types and store BinCut payload voltages into the siteDouble array.
'20200710: Modified to merge the redundant site-loop for voltage calculation.
'20200130: Modified to store non-binning CorePower and OtherRail voltages in the siteDouble array "BinCut_Payload_Voltage()".
'20191202: Modified for the revised initVddBinCondition.
'20191127: Modified for the revised InitVddBinTable.
'20190612: Modified to replace the siteVariables vdd_xxx_val with the new siteVariable bincut_voltage(powerPin).
'20180716: Modified for BinCut testjob mapping.
Public Function bincut_power_Setting_VT(inst_info As Instance_Info, PassBinNum As SiteLong, ByRef voltage_forBinCut() As SiteDouble, Optional enable_DynamicOffset As Boolean = False)
    Dim i As Long
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. BinCut_Payload_Voltage   : siteDouble array to store BinCut payload voltage applied to DCVS.
'''2. Only BinCut LVCC test instances(BV instances, not PostBinCut or OutsideBinCut) have to calculate BinCut payload voltages with DynamicOffset.
'''//==================================================================================================================================================================================//'''
    If UBound(voltage_forBinCut) = MaxBincutPowerdomainCount Then
        '''================= OtherRail setting =======================================
        For i = 0 To UBound(pinGroup_OtherRail) '''OtherRail powerDomains are listed in FullOtherRailinFlowSheet
            Call OtherRail_Pwr_Setting(inst_info, VddBinStr2Enum(pinGroup_OtherRail(i)), PassBinNum, voltage_forBinCut(VddBinStr2Enum(pinGroup_OtherRail(i))))
        Next i
        
        '''================= CorePower setting =======================================
        '''//Apply_testcondition_InFlowSheet() is used to calculate voltage only for non-binning CorePower.
        '''//Decide powerDomain of the binning performance mode, then calculate voltages all non-binning CorePower.
        For i = 0 To UBound(pinGroup_CorePower) '''//CorePower powerDomains are listed in FullCorePowerinFlowSheet.
            '''20210824: Modified to remove the vbt code of checking the powerDomain of the binning performance mode.
            Call Apply_testcondition_InFlowSheet(inst_info, VddBinStr2Enum(pinGroup_CorePower(i)), PassBinNum, voltage_forBinCut(VddBinStr2Enum(pinGroup_CorePower(i))))
        Next i
        
        '''//Calculate BinCut payload voltage with dynamic offset for the binning corePower.
        '''Note: Only BinCut LVCC test instances(BV instances, not PostBinCut or OutsideBinCut) have to calculate BinCut payload voltages with DynamicOffset.
        '''20210824: Modified to move the vbt function Calculate_Binning_CorePower_with_DynamicOffset from calculate_payload_voltage_for_BV to GradeSearch_VT.
        '''20210908: Modified to add the argument "Optional enable_DynamicOffset As Boolean = False" to calculate BinCut payload voltage of the binning PowerDomain with DynamicOffset.
        If enable_DynamicOffset = True Then
            Call Calculate_Binning_CorePower_with_DynamicOffset(inst_info, PassBinNum, voltage_forBinCut(VddBinStr2Enum(inst_info.powerDomain)))
        End If
        
        If inst_info.is_DevChar_Running = True Then
            '''//Set BinCut Payload voltages from DevChar testConditions.
            Call set_BinCut_voltage_for_DevChar(voltage_forBinCut)
        End If
    Else
        TheExec.Datalog.WriteComment "Array size of the argument voltage_forBinCut for bincut_power_Setting_VT doesn't match number of bincut powerDomain MaxBincutPowerdomainCount. Error!!!"
        TheExec.ErrorLogMessage "Array size of the argument voltage_forBinCut for bincut_power_Setting_VT doesn't match number of bincut powerDomain MaxBincutPowerdomainCount. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of bincut_power_Setting_VT"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210914: Modified to revise the vbt code for testCondition such as "cp_gb_hot" and "ft_gb_hot".
'20210824: Modified to add the branch to calculate BinCut payload voltages for testCondition with keyword "*evaluate*bin*".
'20210823: Modified to revise the vbt code for testCondition such as "950mV (ms004 Product)", "MAX(MD003 E1 Product +10%, VDD_SRAM_SOC CPHV)".
'20210207: Modified to print powerDomain of testCondition for the error message.
'20201116: Modified to use "testCondition = LCase(Trim())".
'20200106: Modified to remove the ErrorLogMessage.
'20190527: Modified for C651 new string formats, ex: "MCS601 CPVmax", "MCS608 CPHV", "MCS601 Product-CP2GB", "MCS608 QAHV", "CPVmax", "Product-CP2GB".
'20190325: As CBCHENI's suggestion, we renamed the function from "Calculate_Additional_Voltage_for_PowerRail" to "Calculate_Extra_Voltage_for_PowerRail".
'20190304: Modified for the new string "MS001 Evaluate Bin +15%".
Public Function Calculate_Extra_Voltage_for_PowerRail(testCondition As String, ByRef voltage_val As Double)
    Dim split_content() As String
    Dim strExtraVoltage As String
    Dim str_testCondition As String
On Error GoTo errHandler
    '''//Make sure that the input string in low case.
    str_testCondition = LCase(Trim(testCondition))
    
    '''If the sheet "Non_Binning_Rail" comes with the new keyword, please add the branch for it!!!
    If str_testCondition Like "*h*l*a*m*" Then
        split_content = Split(str_testCondition, " ")
    '''20210823: Modified to revise the vbt code for testCondition such as "950mV (ms004 Product)", "MAX(MD003 E1 Product +10%, VDD_SRAM_SOC CPHV)".
    ElseIf str_testCondition Like "*(*)*" Then '''ex: "950mV (ms004 Product)", "MAX(MD003 E1 Product +10%, VDD_SRAM_SOC CPHV)"
        split_content = Split(str_testCondition, ")")
    ElseIf str_testCondition Like "*vmin*" Then '''ex: "CPVmin"
        split_content = Split(str_testCondition, "vmin")
    ElseIf str_testCondition Like "*vmax*" Then '''ex: "CPVmax"
        split_content = Split(str_testCondition, "vmax")
    ElseIf str_testCondition Like "*lvcc*" Then
        split_content = Split(str_testCondition, "lvcc")
    ElseIf str_testCondition Like "*lv*" Then
        split_content = Split(str_testCondition, "lv")
    ElseIf str_testCondition Like "*hvcc*" Then
        split_content = Split(str_testCondition, "hvcc")
    ElseIf str_testCondition Like "*hv*" Then
        split_content = Split(str_testCondition, "hv")
    ElseIf str_testCondition Like "m*e*voltage*" Then
        split_content = Split(str_testCondition, "voltage")
    ElseIf str_testCondition Like "m*evaluate*bin*" Then
        split_content = Split(str_testCondition, "bin")
    ElseIf str_testCondition Like "m*bin*result*" Then
        split_content = Split(str_testCondition, "result")
    ElseIf str_testCondition Like "*product*-*gb*" Then '''ex: "MI001 Product-CP2GB", "MCS603 Product-CP2GB", "Product-CP2GB".
        '''20210914: Modified to revise the vbt code for testCondition such as "cp_gb_hot" and "ft_gb_hot".
        If str_testCondition Like "*_gb_room*" Then '''ex: "ft_gb_room".
            split_content = Split(str_testCondition, "room")
        ElseIf str_testCondition Like "*_gb_hot*" Then '''ex: "cp_gb_hot", "ft_gb_hot".
            split_content = Split(str_testCondition, "hot")
        ElseIf str_testCondition Like "*-*cp*gb*" Or str_testCondition Like "*-*qa*gb*" Then
            split_content = Split(str_testCondition, "gb")
        Else
            TheExec.Datalog.WriteComment "Non_Binning_Rail: " & testCondition & ". It has the keyword error about calculating the extra voltage! Error!!!"
            'TheExec.ErrorLogMessage "Non_Binning_Rail: " & testCondition & ". It has the keyword error about calculating the extra voltage! Error!!!"
        End If
    ElseIf str_testCondition Like "*product*" Then
        split_content = Split(str_testCondition, "product")
    ElseIf str_testCondition Like "*mv*" Then
        split_content = Split(str_testCondition, "mv")
    Else
        TheExec.Datalog.WriteComment "testCondition:" & testCondition & " from sheet Non_Binning_Rail has the keyword error about calculating the extra voltage! Error!!!"
        'TheExec.ErrorLogMessage "testCondition:" & testCondition & " from sheet Non_Binning_Rail has the keyword error about calculating the extra voltage! Error!!!"
    End If
    
    '''//Calculat the extra voltage
    If Trim(split_content(UBound(split_content))) <> "" Then
        If Trim(split_content(UBound(split_content))) Like "*mv*" Then 'ex: "MS003 bin result -10mv".
            strExtraVoltage = Replace(Trim(split_content(UBound(split_content))), "mv", "")
            voltage_val = voltage_val + CDbl(strExtraVoltage)
        ElseIf Trim(split_content(UBound(split_content))) Like "*%*" Then 'ex: "MS003 bin result -6%".
            strExtraVoltage = Replace(Trim(split_content(UBound(split_content))), "%", "")
            voltage_val = voltage_val * (1 + CDbl(strExtraVoltage) / 100)
        Else
            TheExec.Datalog.WriteComment "Non_Binning_Rail has the incorrect testCondition:" & testCondition & ". Error occurs in Calculate_Extra_Voltage_for_PowerRail. Error!!!"
            'TheExec.ErrorLogMessage "Non_Binning_Rail has the incorrect testCondition:" & testCondition & ". Error occurs in Calculate_Extra_Voltage_for_PowerRail. Error!!!"
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Calculate_Extra_Voltage_for_PowerRail"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Calculate_Extra_Voltage_for_PowerRail"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210621: Modified to remove testLimit and use the fail-stop flag here.
'20201118: Modified to merge the branches.
'20201112: Modified to use "LCase(split_keyword)".
'20200106: Modified to remove the ErrorLogMessage.
'20190304: Modified for splitting the string "MAX(MD003 Product +10%, VDD_SRAM_SOC CP HVCC)".
Public Function Split_special_string_for_calc(str_Input As String, split_keyword As String, str_output() As String)
    Dim strPwrbinVoltage As String
    Dim thissite As Long
On Error GoTo errHandler
    If str_Input <> "" And split_keyword <> "" Then
        '''ex: "MAX(MD003 Product +10%, VDD_SRAM_SOC CP HVCC)"
        strPwrbinVoltage = Replace(LCase(Trim(str_Input)), LCase(split_keyword), "")
        strPwrbinVoltage = Replace(Trim(strPwrbinVoltage), "(", "")
        strPwrbinVoltage = Replace(Trim(strPwrbinVoltage), ")", "")
        str_output = Split(strPwrbinVoltage, ",")
    Else
        '''20210621: Modified to remove testLimit and use the fail-stop flag here.
        thissite = TheExec.sites.SiteNumber
        TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
        TheExec.Datalog.WriteComment "site:" & thissite & ",str_Input:" & str_Input & ",split_keyword:" & split_keyword & ", it has the error while Split_special_string_for_calc. Error!!!"
        'TheExec.ErrorLogMessage "site:" & thissite & ",str_Input:" & str_Input & ",split_keyword:" & split_keyword & ", it has the error while Split_special_string_for_calc. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Split_special_string_for_calc"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Split_special_string_for_calc"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210917: Modified to revise the branch for testCondition such as "ME005 BinX Product + 10%", as requested by TSMC ZYLINI.
'20210906: Modified to remove the redundant argument "ByRef voltage_SelsrmBitCalc As SiteDouble".
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210824: Modified to add the branch to calculate BinCut payload voltages for testCondition with keyword "*evaluate*bin*".
'20210823: Modified to revise the vbt code for C651 new rules of Guardband naming, requested by C651 Toby and TSMC ZYLINI.
'20210820: Modified to revise the vbt code for testCondition with the new format such as "MG008 BinX Product +10%".
'20210819: Modified to get the performance mode from testCondition such as "950mV (MS004 Product)", as requested by TSMC ZYLINI.
'20210817: Modified to check the keyword in testCondition for BinCut voltage calculation.
'20210817: Modified to get additional condition from testCondition of powerDomain.
'20210813: Modified to adjust the priority of parsing testcondition, as requested by TSMC ZYLINI.
'20210812: Modified to remove checking "If AllBinCut(nonbinning_pmode).is_for_BinSearch = True" for testCondition with keyword "m*bin*result*".
'20210812: As per discussion with C651 Toby and ZYLINI, since HVCC tests has "m*##*Bin*Product" or "m*##*E1*Product" with p_mode not tested, we have to create the special case for this.
'20210803: Modified to check if bincutNum>PassBinCut_ary(Ubound(PassBinCut_ary)).
'20210802: Modified to replace is_BinCutJob_for_StepSearch with AllBinCut(p_mode).is_for_BinSearch or inst_info.is_BinSearch.
'20210802: Modified to check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY)
'20210802: Modified to merge the vbt functions Calculate_LVCC_for_CorePower and Calculate_HVCC_for_CorePower into the vbt function Calculate_Voltage_for_CorePower.
Public Function Calculate_Voltage_for_CorePower(inst_info As Instance_Info, idxPin_CorePower As Integer, PassBinNum As SiteLong, testCondition As String, voltage_val As SiteDouble)
    Dim thissite As Long
    Dim split_content() As String
    Dim nonbinning_pmode As Integer
    Dim strTemp As String
    Dim str_additional_condition As String
    Dim EQ_Num As Long
    Dim step_in_DynamicIdsZone As Long
    Dim voltage_Temp As Double
    Dim CP_Vmax As Double
    Dim CP_Vmin As Double
    Dim c As Double
    Dim M As Double
    Dim bincutNum As Long
    Dim str_testCondition As String
    Dim is_Selsrm_LogicPower As Boolean
    Dim idx_sramVth As Long
    Dim performance_mode As String
    Dim dbl_GB_BinCutJob As Double
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''20210728: C651 Toby revised the rules about SRAMthreshold selection for SELSRM bit calculation.
'''1) For the testCondition with keyword "M*### E1 Voltage*", SELSRM bit calculation uses EQN-based voltage for SELSRM logic power and SRAMthersh_binsearch for SRAM_Vth.
'''2) For the testCondition with keyword "M*### Bin Result*", SELSRM bit calculation uses EQN-based voltage for SELSRM logic power and SRAMthersh_binsearch for SRAM_Vth.
'''3) For the testCondition with keyword "M*### Product*", SELSRM bit calculation uses Efuse product voltage for SELSRM logic power and SRAMthersh_product for SRAM_Vth.
'''4) For the testCondition with keyword "MAX(*,*)", SELSRM bit calculation uses the max value for SELSRM logic power and SRAMthersh_product for SRAM_Vth.
'''20210812: As per discussion with C651 Toby and ZYLINI, since HVCC tests has "m*##*Bin*Product" or "m*##*E1*Product" with p_mode not tested, we have to create the special case for this.
'''//==================================================================================================================================================================================//'''
    '''init
    thissite = TheExec.sites.SiteNumber
    str_testCondition = LCase(Trim(testCondition))
    performance_mode = ""
    
    '''//Check if powerDomain is the SELSRM logic power.
    If LCase("*," & selsramLogicPin & ",*") Like LCase("*," & VddBinName(idxPin_CorePower) & ",*") Then
        is_Selsrm_LogicPower = True
        
        If str_testCondition Like "*product*" Then
            idx_sramVth = 1 '''SRAMthersh_product
        Else
            idx_sramVth = 0 '''SRAMthersh_binsearch
        End If
    Else
        is_Selsrm_LogicPower = False
    End If
    
    '''//Check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY)...
    '''20210802: Modified to check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY).
    If str_testCondition Like "*bin1*" Then '''Bin1
        bincutNum = 1
    ElseIf str_testCondition Like "*binx*" Then '''BinX
        bincutNum = 2
    ElseIf str_testCondition Like "*biny*" Then '''BinY
        bincutNum = 3
    Else '''If testCondition doesn't contain any keyword about PassBin, use PassBinNum(from the current step of the binning p_mode).
        bincutNum = PassBinNum
    End If
    
    '''20210803: Modified to check if bincutNum>PassBinCut_ary(Ubound(PassBinCut_ary)).
    If bincutNum > PassBinCut_ary(UBound(PassBinCut_ary)) Then
        TheExec.sites.Item(thissite).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
        TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition: " & testCondition & ", it has the incorrect keyword in testCondition about the BinCut number for Calculate_Voltage_for_CorePower. Error!!!"
        TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition: " & testCondition & ", it has the incorrect keyword in testCondition about the BinCut number for Calculate_Voltage_for_CorePower. Error!!!"
        Exit Function
    End If
    
    '''//Check keyword in the testCondition to calculate BinCut payload voltage.
    If str_testCondition Like "m*##*" Then 'ex: "MC601 E1 Voltage".
        '''//Split string of the testCondition to get performance mode of CorePower.
        split_content = Split(str_testCondition, " ")
        performance_mode = UCase(Trim(split_content(0)))
        
        '''//Check if any correct keyword of performance_mode exists...
        If VddbinPmodeDict.Exists(performance_mode) Then
            '''//p_mode of non_binning CorePower.
            nonbinning_pmode = VddBinStr2Enum(performance_mode)
            '''20210817: Modified to get additional condition from testCondition of powerDomain.
            str_additional_condition = LCase(Trim(Replace(str_testCondition, LCase(performance_mode), "")))
        Else
            str_additional_condition = ""
            TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", keyword:" & performance_mode & " is not any correct performance mode defined in BinCut voltage table. Please check the testCondition in sheet Non_Binning_Rail for Calculate_Voltage_for_CorePower. Error!!!"
            TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", keyword:" & performance_mode & " is not any correct performance mode defined in BinCut voltage table. Please check the testCondition in sheet Non_Binning_Rail for Calculate_Voltage_for_CorePower. Error!!!"
        End If
        
        '''20210824: Modified to add the branch to calculate BinCut payload voltages for testCondition with keyword "*evaluate*bin*".
        If str_additional_condition Like "*evaluate*bin*" Then '''ex: "MS001 Evaluate Bin +6%".
            If AllBinCut(nonbinning_pmode).powerPin = inst_info.powerDomain Then
                voltage_Temp = DYNAMIC_VBIN_IDS_ZONE(nonbinning_pmode).Voltage(inst_info.step_Current(thissite))
                
                '''//for SELSRM DSSC bit calculation
                If is_Selsrm_LogicPower = True Then
                    '''20210728: C651 Toby revised the rules about SRAMthreshold selection for SELSRM bit calculation.
                    inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = voltage_Temp '''EQN-based LVCC.
                    inst_info.sram_Vth(idxPin_CorePower)(thissite) = Floor(BinCut(nonbinning_pmode, bincutNum).SRAM_VTH_SPEC(idx_sramVth))
                End If
            Else
                voltage_Temp = 0
                inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = 0
                TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", keyword:" & performance_mode & " doesn't belong to the same powerDomain of " & inst_info.performance_mode & ". Please check the testCondition in sheet Non_Binning_Rail for Calculate_Voltage_for_CorePower. Error!!!"
                TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", keyword:" & performance_mode & " doesn't belong to the same powerDomain of " & inst_info.performance_mode & ". Please check the testCondition in sheet Non_Binning_Rail for Calculate_Voltage_for_CorePower. Error!!!"
            End If
            
        '''20210817: Modified to check the keyword in testCondition for BinCut voltage calculation.
        ElseIf str_additional_condition Like "*bin*result*" Then '''ex: "MS003 bin result +6%", "MS003 bin result +10mv".
            '''//Check if p_mode is searched.
            If VBIN_RESULT(nonbinning_pmode).tested = True Then
                voltage_Temp = VBIN_RESULT(nonbinning_pmode).GRADE
                
                '''//for SELSRM DSSC bit calculation
                If is_Selsrm_LogicPower = True Then
                    '''20210728: C651 Toby revised the rules about SRAMthreshold selection for SELSRM bit calculation.
                    inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = voltage_Temp '''EQN-based LVCC.
                    inst_info.sram_Vth(idxPin_CorePower)(thissite) = Floor(BinCut(nonbinning_pmode, bincutNum).SRAM_VTH_SPEC(idx_sramVth))
                End If
                
            ElseIf VBIN_RESULT(nonbinning_pmode).tested = False And inst_info.is_DevChar_Running = True Then
                '''//Get step in Dynamic IDS Zone for nonbinning_pmode.
                '''//If p_mode is not tested, assign EQN1 voltage for DevChar, requested by PCLIN.
                '''20210302: Modified to use EQN1 to calculate voltage for "M*### Bin Result" if M*### is not tested.
                bincutNum = CurrentPassBinCutNum(thissite)
                step_in_DynamicIdsZone = DYNAMIC_VBIN_IDS_ZONE(nonbinning_pmode).step_Mapping(bincutNum, 1)
                
                '''//Get BinCut voltages by step in Dynamic IDS Zone for nonbinning_pmode.
                If step_in_DynamicIdsZone > -1 Then
                    voltage_Temp = DYNAMIC_VBIN_IDS_ZONE(nonbinning_pmode).Voltage(step_in_DynamicIdsZone)
                    
                    '''//for SELSRM DSSC bit calculation
                    If is_Selsrm_LogicPower = True Then
                        '''20210728: C651 Toby revised the rules about SRAMthreshold selection for SELSRM bit calculation.
                        inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = voltage_Temp '''EQN-based LVCC.
                        inst_info.sram_Vth(idxPin_CorePower)(thissite) = Floor(BinCut(nonbinning_pmode, bincutNum).SRAM_VTH_SPEC(idx_sramVth))
                    End If
                Else
                    voltage_Temp = 0
                    inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = 0
                    TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", it doesn't have any correct step in DYNAMIC_IDS_Zone for Bin=" & bincutNum & " and EQN=1. Please check Dev_Char settings for Calculate_Voltage_for_CorePower. Error!!!"
                    TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", it doesn't have any correct step in DYNAMIC_IDS_Zone for Bin=" & bincutNum & " and EQN=1. Please check Dev_Char seetings for Calculate_Voltage_for_CorePower. Error!!!"
                End If
            Else
                TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", maybe it has the incorrect keyword 'bin result', or " & VddBinName(nonbinning_pmode) & " is not tested. Please check the testCondition in sheet Non_Binning_Rail for Calculate_Voltage_for_CorePower. Error!!!"
                TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", maybe it has the incorrect keyword 'bin result', or " & VddBinName(nonbinning_pmode) & " is not tested. Please check the testCondition in sheet Non_Binning_Rail for Calculate_Voltage_for_CorePower. Error!!!"
            End If
            
        ElseIf str_additional_condition Like "*e#*" Then '''ex: "MC601 E1 Voltage", "MC601 E1 Product".
            '''//Get BinCut EQN
            strTemp = split_content(1)
            split_content = Split(Trim(strTemp), "e")
            EQ_Num = CLng(split_content(1))
            
            '''//Get step in Dynamic IDS Zone for nonbinning_pmode.
            step_in_DynamicIdsZone = DYNAMIC_VBIN_IDS_ZONE(nonbinning_pmode).step_Mapping(bincutNum, EQ_Num)
            
            '''//Get BinCut voltages by step in Dynamic IDS Zone for nonbinning_pmode.
            If step_in_DynamicIdsZone > -1 Then
                '''//Check if tesCondition contains any keyword about "Voltage", "Product".
                If str_testCondition Like "m*e*voltage*" Then '''ex: "MC601 E1 Voltage"
                    voltage_Temp = DYNAMIC_VBIN_IDS_ZONE(nonbinning_pmode).Voltage(step_in_DynamicIdsZone)
                ElseIf str_testCondition Like "m*e*product*" Then '''ex: "MC601 Product".
                    voltage_Temp = DYNAMIC_VBIN_IDS_ZONE(nonbinning_pmode).Product_Voltage(step_in_DynamicIdsZone)
                Else
                    voltage_Temp = 0
                    inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = 0
                    TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", it doesn't have any correct keyword for Calculate_LVCC_for_CorePower. Please check the testCondition in sheet Non_Binning_Rail. Error!!!"
                    TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", it doesn't have any correct keyword for Calculate_LVCC_for_CorePower. Please check the testCondition in sheet Non_Binning_Rail. Error!!!"
                End If
                
                '''//for SELSRM DSSC bit calculation
                If is_Selsrm_LogicPower = True Then
                    '''20210728: C651 Toby revised the rules about SRAMthreshold selection for SELSRM bit calculation.
                    inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = voltage_Temp '''EQN-based voltage.
                    inst_info.sram_Vth(idxPin_CorePower)(thissite) = Floor(BinCut(nonbinning_pmode, bincutNum).SRAM_VTH_SPEC(idx_sramVth))
                End If
            Else
                voltage_Temp = 0
                inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = 0
                TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", it doesn't have any correct step in DYNAMIC_IDS_Zone for Calculate_LVCC_for_CorePower. Please check the testCondition in sheet Non_Binning_Rail. Error!!!"
                TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", it doesn't have any correct step in DYNAMIC_IDS_Zone for Calculate_LVCC_for_CorePower. Please check the testCondition in sheet Non_Binning_Rail. Error!!!"
            End If

        ElseIf str_additional_condition Like "*product*" Then '''ex: "MC601 Product", "ME005 BinX Product + 10%"
            '''20210812: As per discussion with C651 Toby and ZYLINI, since HVCC tests has "m*##*Bin*Product" or "m*##*E1*Product" with p_mode not tested, we have to create the special case for this.
            '''20210820: Modified to revise the vbt code for testCondition with the new format such as "MG008 BinX Product +10%".
            '''20210917: Modified to revise the branch for testCondition such as "ME005 BinX Product + 10%", as requested by TSMC ZYLINI.
            If str_additional_condition Like "*bin*" Then '''ex: "ME005 BinX Product + 10%"
                '''//Get step in Dynamic IDS Zone for nonbinning_pmode.
                step_in_DynamicIdsZone = DYNAMIC_VBIN_IDS_ZONE(nonbinning_pmode).step_Mapping(bincutNum, 1)
                
                '''//Get BinCut voltages by step in Dynamic IDS Zone for nonbinning_pmode.
                If step_in_DynamicIdsZone > -1 Then
                    voltage_Temp = DYNAMIC_VBIN_IDS_ZONE(nonbinning_pmode).Product_Voltage(step_in_DynamicIdsZone)
                    
                    '''//for SELSRM DSSC bit calculation
                    If is_Selsrm_LogicPower = True Then
                        '''20210728: C651 Toby revised the rules about SRAMthreshold selection for SELSRM bit calculation.
                        inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = voltage_Temp '''product voltage.
                        inst_info.sram_Vth(idxPin_CorePower)(thissite) = Floor(BinCut(nonbinning_pmode, bincutNum).SRAM_VTH_SPEC(idx_sramVth))
                    End If
                Else
                    voltage_Temp = 0
                    inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = 0
                    TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", " & VddBinName(nonbinning_pmode) & " is not tested, but no step in Dynamic_IDS_Zone to calculate product voltage for Calculate_Voltage_for_CorePower. Error!!!"
                    TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", " & VddBinName(nonbinning_pmode) & " is not tested, but no step in Dynamic_IDS_Zone to calculate product voltage for Calculate_Voltage_for_CorePower. Error!!!"
                End If
            Else
                If VBIN_RESULT(nonbinning_pmode).tested = True Then
                    voltage_Temp = VBIN_RESULT(nonbinning_pmode).GRADEVDD
                    
                    '''//for SELSRM DSSC bit calculation
                    If is_Selsrm_LogicPower = True Then
                        '''20210728: C651 Toby revised the rules about SRAMthreshold selection for SELSRM bit calculation.
                        inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = voltage_Temp '''product voltage.
                        inst_info.sram_Vth(idxPin_CorePower)(thissite) = Floor(BinCut(nonbinning_pmode, bincutNum).SRAM_VTH_SPEC(idx_sramVth))
                    End If
                Else
                    voltage_Temp = 0
                    inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = 0
                    TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", " & VddBinName(nonbinning_pmode) & " is not tested, but keyword 'product' requests to calculate product voltage for Calculate_Voltage_for_CorePower. Error!!!"
                    TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", " & VddBinName(nonbinning_pmode) & " is not tested, but keyword 'product' requests to calculate product voltage for Calculate_Voltage_for_CorePower. Error!!!"
                End If
            End If
        Else
            TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", it doesn't have any correct keyword about " & nonbinning_pmode & " for Calculate_Voltage_for_CorePower. Error!!!"
            'TheExec.ErrorLogMessage "site:" & thissite & "," &  inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", it doesn't have any correct keyword about " & nonbinning_pmode & " for Calculate_Voltage_for_CorePower. Error!!!"
        End If
        
    ElseIf str_testCondition Like "#*mv*" Then '''ex: "650mV"
        '''//Check if testCondition contains any keyword about "###mv"
        split_content = Split(str_testCondition, "mv")
        voltage_Temp = CDbl(LCase(Trim(split_content(0))))
            
        '''//for SELSRM DSSC bit calculation
        If is_Selsrm_LogicPower = True Then
            '''20210728: C651 Toby revised the rules about SRAMthreshold selection for SELSRM bit calculation.
            inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = voltage_Temp
            
            If Trim(split_content(1)) Like "*(*)*" Then '''ex: "900mV (MS003)", "950mV (MS004 Product)"
                strTemp = Replace(LCase(Trim(split_content(1))), "(", "")
                strTemp = Replace(Trim(strTemp), ")", "")
                '''20210819: Modified to get the performance mode from testCondition such as "950mV (MS004 Product)", as requested by TSMC ZYLINI.
                split_content = Split(strTemp, " ")
                performance_mode = UCase(Trim(split_content(0)))
                                
                '''//Check if performance_mode is defined in BinCut voltage tables.
                If VddbinPmodeDict.Exists(performance_mode) Then
                    nonbinning_pmode = VddBinStr2Enum(UCase(Trim(performance_mode)))
                    inst_info.sram_Vth(idxPin_CorePower)(thissite) = Floor(BinCut(nonbinning_pmode, bincutNum).SRAM_VTH_SPEC(idx_sramVth))
                Else
                    inst_info.sram_Vth(idxPin_CorePower)(thissite) = 0
                    TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", keyword:" & performance_mode & " is not defined in any BinCut voltage table. Please check the testCondition in sheet Non_Binning_Rail. Error!!!"
                    TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", keyword:" & performance_mode & " is not defined in any BinCut voltage table. Please check the testCondition in sheet Non_Binning_Rail. Error!!!"
                End If
            Else
                inst_info.sram_Vth(idxPin_CorePower)(thissite) = 0
                TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", it doesn't have any correct keyword about p_mode to get SRAM_Vth for Calculate_Voltage_for_CorePower. Please check the testCondition in sheet Non_Binning_Rail. Error!!!"
                TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition:" & testCondition & ", it doesn't have any correct keyword about p_mode to get SRAM_Vth for Calculate_Voltage_for_CorePower. Please check the testCondition in sheet Non_Binning_Rail. Error!!!"
            End If
        End If
    Else
        voltage_Temp = 0
        inst_info.voltage_SelsrmBitCalc(idxPin_CorePower)(thissite) = 0
        TheExec.sites.Item(thissite).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
        TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition: " & testCondition & ", it doesn't have any keyword in testCondition about payload voltage for Calculate_Voltage_for_CorePower. Error!!!"
        TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition: " & testCondition & ", it doesn't have any keyword in testCondition about payload voltage for Calculate_Voltage_for_CorePower. Error!!!"
        Exit Function
    End If
        
    '''//Find GuardBand(GB) by the keyword in the testCondition.
    '''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
    '''20210823: Modified to revise the vbt code for C651 new rules of Guardband naming, requested by C651 Toby and TSMC ZYLINI.
    If str_testCondition Like "*gb*" Then
        If str_testCondition Like LCase("*CPGB*") Or str_testCondition Like LCase("*BinningGB*") Then
            dbl_GB_BinCutJob = BinCut(nonbinning_pmode, bincutNum).CP_GB(0)
        ElseIf str_testCondition Like LCase("*CP2GB*") Or str_testCondition Like LCase("*CP_GB_HOT*") Then
            dbl_GB_BinCutJob = BinCut(nonbinning_pmode, bincutNum).CP2_GB(0)
        ElseIf str_testCondition Like LCase("*FT_GB_ROOM*") Then
            dbl_GB_BinCutJob = BinCut(nonbinning_pmode, bincutNum).FT1_GB(0)
        ElseIf str_testCondition Like LCase("*FT_GB_HOT*") Then
            dbl_GB_BinCutJob = BinCut(nonbinning_pmode, bincutNum).FT2_GB(0)
        ElseIf str_testCondition Like LCase("*SLTGB*") Then
            dbl_GB_BinCutJob = BinCut(nonbinning_pmode, bincutNum).SLT_GB(0)
        ElseIf str_testCondition Like LCase("*ATE_FQAGB*") Then
            dbl_GB_BinCutJob = BinCut(nonbinning_pmode, bincutNum).FTQA_GB(0)
        ElseIf str_testCondition Like LCase("*HTOL_RO_GB*") Then
            dbl_GB_BinCutJob = BinCut(nonbinning_pmode, bincutNum).HTOL_RO_GB(0)
        ElseIf str_testCondition Like LCase("*HTOL_RO_GB_ROOM*") Then
            dbl_GB_BinCutJob = BinCut(nonbinning_pmode, bincutNum).HTOL_RO_GB_ROOM(0)
        ElseIf str_testCondition Like LCase("*HTOL_RO_GB_HOT*") Then
            dbl_GB_BinCutJob = BinCut(nonbinning_pmode, bincutNum).HTOL_RO_GB_HOT(0)
        ElseIf str_testCondition Like LCase("*SLT_FQA_GB*") Then
            dbl_GB_BinCutJob = BinCut(nonbinning_pmode, bincutNum).SLT_FTQA_GB(0)
        Else
            dbl_GB_BinCutJob = 0
            TheExec.sites.Item(thissite).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
            TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition: " & testCondition & ", it doesn't have any keyword in testCondition to find the matched guardBand for Calculate_Voltage_for_CorePower. Error!!!"
            TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_CorePower) & ",testCondition: " & testCondition & ", it doesn't have any keyword in testCondition to find the matched guardBand for Calculate_Voltage_for_CorePower. Error!!!"
        End If
        
        '''//Check if testCondition contains any keyword about "*GB*".
        If str_testCondition Like "*+*gb*" Then
            voltage_Temp = voltage_Temp + dbl_GB_BinCutJob
        ElseIf str_testCondition Like "*-*gb*" Then 'ex: "MC601 Product-CP2GB"
            voltage_Temp = voltage_Temp - dbl_GB_BinCutJob
        End If
    End If
    
    '''//Check if testCondition with "mv" or "%", ex: "+6%" of "MS003 product +6%".
    Calculate_Extra_Voltage_for_PowerRail testCondition, voltage_Temp
    
    '''//Output the voltage value.
    voltage_val(thissite) = voltage_Temp
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Calculate_Voltage_for_CorePower"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210906: Modified to remove the redundant argument "ByRef voltage_SelsrmBitCalc As SiteDouble".
'20210823: Modified to revise the vbt code for C651 new rules of Guardband naming, requested by C651 Toby and TSMC ZYLINI.
'20210820: Modified to revised the vbt code to support the new format of testCondition such as "MPS001 BinningVmax", requested by TSMC ZYLINI.
'20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'20210803: Modified to check if bincutNum>PassBinCut_ary(Ubound(PassBinCut_ary)).
'20210802: Modified to check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY).
'20210703: Modified to revise the error message about the incorrect keyword in the testCondition of BinCut powerDomain.
'20210623: Modified to revise the vbt function adjust_VddBinning for BinCut search in FT.
'20210207: Modified to print powerDomain of testCondition for the error message.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20201116: Modified to use "testCondition = LCase(Trim())".
'20201029: Modified to replace the argument "other_rail_mode as Integer" with "idxPin_OtherRail As Integer".
'20201029: Modified to use "Public Type Instance_Info".
'20201013: Modified to check if performance mode exists.
'20201013: Modified to revise the methods of parsing testConditions and calculating BinCut payload voltages.
'20200812: Modified to replace "other_rail_val_site" with "voltage_OtherRail".
'20200806: Modified to use siteDouble for Calculate_LVCC_for_CorePower, Calculate_HVCC_for_CorePower, and Calculate_Voltage_for_OtherRail.
'20200727: Modified to use "voltage_SelsrmBitCalc" to store voltage for SELSRM bit calculation.
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20191127: Modified for the revised InitVddBinTable.
'20190605: Modified for OtherRail with CorePower p_mode.
'20190527: Modified for C651 new string formats, ex: "MCS601 CPVmax", "MCS608 CPHV", "MCS601 Product-CP2GB", "MCS608 QAHV", "CPVmax", "Product-CP2GB".
'20190315: Created for calculating OtherRail LVCC/HVCC voltage.
Public Function Calculate_Voltage_for_OtherRail(inst_info As Instance_Info, idxPin_OtherRail As Integer, PassBinNum As SiteLong, testCondition As String, voltage_OtherRail As SiteDouble)
    Dim thissite As Long
    Dim split_content() As String
    Dim voltage_Temp As Double
    Dim power_GB As Double
    Dim performance_mode As String
    Dim main_power_mode As Integer
    Dim is_Keyword_correct As Boolean
    Dim str_testCondition As String
    Dim bincutNum As Long
    Dim dbl_GB_BinCutJob As Double
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Warning!!!
'''ToDo: Currently C651 Toby didn't define the format of non_binning_rail for HBV retention and postbincut.
'''So we kept the original non_binning_rail here. 20180921
'''Once if Toby revised the definition, we will modify the following code...
'''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'''//==================================================================================================================================================================================//'''
    '''init
    thissite = TheExec.sites.SiteNumber
    main_power_mode = -1
    is_Keyword_correct = True
    str_testCondition = LCase(Trim(testCondition))
    
    '''//Check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY)...
    '''20210802: Modified to check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY).
    If str_testCondition Like "*bin1*" Then '''Bin1
        bincutNum = 1
    ElseIf str_testCondition Like "*binx*" Then '''BinX
        bincutNum = 2
    ElseIf str_testCondition Like "*biny*" Then '''BinY
        bincutNum = 3
    Else '''If testCondition doesn't contain any keyword about PassBin, use PassBinNum(from the current step of the binning p_mode).
        bincutNum = PassBinNum
    End If
    
    '''20210803: Modified to check if bincutNum>PassBinCut_ary(Ubound(PassBinCut_ary)).
    If bincutNum > PassBinCut_ary(UBound(PassBinCut_ary)) Then
        TheExec.sites.Item(thissite).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
        TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition: " & testCondition & ", it has the incorrect keyword in testCondition about the BinCut number for Calculate_Voltage_for_OtherRail. Error!!!"
        TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition: " & testCondition & ", it has the incorrect keyword in testCondition about the BinCut number for Calculate_Voltage_for_OtherRail. Error!!!"
        Exit Function
    End If
    
    '''//Get the performance mode.
    If str_testCondition Like "m*##*" Then '''ex: "MCS601 CPVmax", "MCS608 CPHV", "MCS601 Product-CP2GB", "MCS608 QAHV".
        '''//Split the string by " ".
        split_content = Split(testCondition, " ")
                
        '''//If SRAM power with CorePower Pmode in Mode column, get main Pmode from the string, ex: "MCS601 CPVmax".
        If Len(Trim(split_content(0))) = 6 Then
            performance_mode = UCase(Trim(Mid(split_content(0), 1, 2) & Mid(split_content(0), 4, 3)))
        Else
            performance_mode = UCase(Trim(split_content(0)))
        End If
        
        '''//Check if performance_mode exists...
        If VddbinPmodeDict.Exists(performance_mode) Then
            main_power_mode = VddBinStr2Enum(performance_mode)
            
            '''//Remove Pmode
            str_testCondition = LCase(Trim(Replace(str_testCondition, split_content(0), "")))
            
            '''//Compatible with string formats of the conventional project and Sicily.
            If str_testCondition Like "*sram*" Then
                split_content = Split(str_testCondition, "sram")
                str_testCondition = Trim(split_content(1))
            End If
            
            '''//If SRAM has keyword "v", we should replace it with "lvcc", ex: "MC001 SRAM V" --> "MC001 SRAM LVCC".
            If testCondition Like "v*" Then
                str_testCondition = Replace(str_testCondition, "v", "lvcc")
            End If
        Else
            TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition:" & testCondition & ", it doesn't have any correct keyword about the performance mode of OtherRail for Calculate_Voltage_for_OtherRail. Error!!!"
            'TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition:" & testCondition & ", it doesn't have any correct keyword about the performance mode of OtherRail for Calculate_Voltage_for_OtherRail. Error!!!"
        End If
    Else
        main_power_mode = inst_info.p_mode
    End If
    
    '''//If performance_mode exists, calculate BinCut payload voltages.
    If main_power_mode <> -1 Then
        '''//Calculate BinCut payload voltages according to keywords from testCondition.
        If str_testCondition Like "*lvcc*" Or str_testCondition Like "*cp*lv*" Then
            If bincutJobName = "cp1" Then
                voltage_Temp = BinCut(main_power_mode, bincutNum).OTHER_CP1_RAIL(idxPin_OtherRail)
            Else
                '''//According to BinCut testJob, select the matched GuardBand.
                Select Case LCase(bincutJobName)
                    Case "cp1": power_GB = BinCut(main_power_mode, bincutNum).OTHER_CP1_GB(idxPin_OtherRail)
                    Case "cp2": power_GB = BinCut(main_power_mode, bincutNum).OTHER_CP2_GB(idxPin_OtherRail)
                    Case "ft_room": power_GB = BinCut(main_power_mode, bincutNum).OTHER_FT1_GB(idxPin_OtherRail)
                    Case "ft_hot": power_GB = BinCut(main_power_mode, bincutNum).OTHER_FT2_GB(idxPin_OtherRail)
                    Case "qa": power_GB = BinCut(main_power_mode, bincutNum).OTHER_ATE_FQA_GB(idxPin_OtherRail)
                    Case Else:
                            TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition:" & testCondition & ", The selected testJob " & bincutJobName & " doesn't have the matched Guardband for Calculate_Voltage_for_OtherRail. Error!!!"
                            TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition:" & testCondition & ", The selected testJob " & bincutJobName & " doesn't have the matched Guardband for Calculate_Voltage_for_OtherRail. Error!!!"
                End Select
            
                voltage_Temp = BinCut(main_power_mode, bincutNum).OTHER_PRODUCT_RAIL(idxPin_OtherRail) - power_GB
            End If
        
        ElseIf str_testCondition Like "*hvcc*" Or str_testCondition Like "*hv*" Then '''ex: "MCS608 CPHV".
            If testCondition = "*cp*hv*" Then
                voltage_Temp = BinCut(main_power_mode, bincutNum).HVCC_OTHER_CP_RAIL(idxPin_OtherRail)
            ElseIf testCondition = "*ft*hv*" Then
                voltage_Temp = BinCut(main_power_mode, bincutNum).HVCC_OTHER_FT_RAIL(idxPin_OtherRail)
            ElseIf testCondition Like "*qa*hv*" Then
                voltage_Temp = BinCut(main_power_mode, bincutNum).HVCC_OTHER_QA_RAIL(idxPin_OtherRail)
            Else
                Select Case LCase(bincutJobName)
                    Case "cp1": voltage_Temp = BinCut(main_power_mode, bincutNum).HVCC_OTHER_CP_RAIL(idxPin_OtherRail)
                    Case "cp2": voltage_Temp = BinCut(main_power_mode, bincutNum).HVCC_OTHER_CP_RAIL(idxPin_OtherRail)
                    Case "ft_room": voltage_Temp = BinCut(main_power_mode, bincutNum).HVCC_OTHER_FT_RAIL(idxPin_OtherRail)
                    Case "ft_hot": voltage_Temp = BinCut(main_power_mode, bincutNum).HVCC_OTHER_FT_RAIL(idxPin_OtherRail)
                    Case "qa": voltage_Temp = BinCut(main_power_mode, bincutNum).HVCC_OTHER_QA_RAIL(idxPin_OtherRail)
                    Case Else:
                            TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition:" & testCondition & ", The selected testJob " & bincutJobName & " doesn't have the matched HVCC of OtherRail for Calculate_Voltage_for_OtherRail. Error!!!"
                            TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition:" & testCondition & ", The selected testJob " & bincutJobName & " doesn't have the matched HVCC of OtherRail for Calculate_Voltage_for_OtherRail. Error!!!"
                End Select
            End If
        '''20210820: Modified to revised the vbt code to support the new format of testCondition such as "MPS001 BinningVmax", requested by TSMC ZYLINI.
        ElseIf str_testCondition Like "*cp*vmax*" Or str_testCondition Like LCase("*BinningVmax*") Then '''ex: "MCS601 CPVmax", "CPVmax", "BinningVmax", "MPS001 BinningVmax".
            voltage_Temp = BinCut(main_power_mode, bincutNum).OTHER_CP_Vmax(idxPin_OtherRail)
        ElseIf str_testCondition Like "*cp*vmin*" Or str_testCondition Like LCase("*BinningVmin*") Then  '''ex: "MCS601 CPVmin", "CPVmin".
            voltage_Temp = BinCut(main_power_mode, bincutNum).OTHER_CP_Vmin(idxPin_OtherRail)
        ElseIf str_testCondition Like "*product*" Then
            voltage_Temp = BinCut(main_power_mode, bincutNum).OTHER_PRODUCT_RAIL(idxPin_OtherRail)
        ElseIf str_testCondition Like "*mv" Then '''ex: 900mv.
            split_content = Split(str_testCondition, "mv")
            voltage_Temp = CDbl(split_content(0))
        Else
            is_Keyword_correct = False
            voltage_Temp = 0
            TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition:" & testCondition & ", it doesn't have any correct keyword to decide the parsing method in Calculate_Voltage_for_OtherRail. Error!!!"
            TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition:" & testCondition & ", it doesn't have any correct keyword to decide the parsing method in Calculate_Voltage_for_OtherRail. Error!!!"
        End If
    
        '''//Find GuardBand(GB) by the keyword in the testCondition.
        '''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
        '''20210823: Modified to revise the vbt code for C651 new rules of Guardband naming, requested by C651 Toby and TSMC ZYLINI.
        If str_testCondition Like "*gb*" Then
            If str_testCondition Like LCase("*CPGB*") Or str_testCondition Like LCase("*BinningGB*") Then
                dbl_GB_BinCutJob = BinCut(main_power_mode, bincutNum).OTHER_CP1_GB(idxPin_OtherRail)
            ElseIf str_testCondition Like LCase("*CP2GB*") Or str_testCondition Like LCase("*CP_GB_HOT*") Then
                dbl_GB_BinCutJob = BinCut(main_power_mode, bincutNum).OTHER_CP2_GB(idxPin_OtherRail)
            ElseIf str_testCondition Like LCase("*FT_GB_ROOM*") Then
                dbl_GB_BinCutJob = BinCut(main_power_mode, bincutNum).OTHER_FT1_GB(idxPin_OtherRail)
            ElseIf str_testCondition Like LCase("*FT_GB_HOT*") Then
                dbl_GB_BinCutJob = BinCut(main_power_mode, bincutNum).OTHER_FT2_GB(idxPin_OtherRail)
            ElseIf str_testCondition Like LCase("*SLTGB*") Then
                dbl_GB_BinCutJob = BinCut(main_power_mode, bincutNum).OTHER_SLT_GB(idxPin_OtherRail)
            ElseIf str_testCondition Like LCase("*ATE_FQAGB*") Then
                dbl_GB_BinCutJob = BinCut(main_power_mode, bincutNum).OTHER_ATE_FQA_GB(idxPin_OtherRail)
            ElseIf str_testCondition Like LCase("*HTOL_RO_GB*") Then
                dbl_GB_BinCutJob = BinCut(main_power_mode, bincutNum).OTHER_HTOL_RO_GB(idxPin_OtherRail)
            ElseIf str_testCondition Like LCase("*HTOL_RO_GB_ROOM*") Then
                dbl_GB_BinCutJob = BinCut(main_power_mode, bincutNum).OTHER_HTOL_RO_GB_ROOM(idxPin_OtherRail)
            ElseIf str_testCondition Like LCase("*HTOL_RO_GB_HOT*") Then
                dbl_GB_BinCutJob = BinCut(main_power_mode, bincutNum).OTHER_HTOL_RO_GB_HOT(idxPin_OtherRail)
            ElseIf str_testCondition Like LCase("*SLT_FQA_GB*") Then
                dbl_GB_BinCutJob = BinCut(main_power_mode, bincutNum).OTHER_SLT_FQA_GB(idxPin_OtherRail)
            Else
                dbl_GB_BinCutJob = 0
                TheExec.sites.Item(thissite).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                TheExec.Datalog.WriteComment "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition: " & testCondition & ", it doesn't have any keyword in testCondition to find the matched guardBand for Calculate_Voltage_for_CorePower. Error!!!"
                TheExec.ErrorLogMessage "site:" & thissite & "," & inst_info.performance_mode & ",powerDomain:" & VddBinName(idxPin_OtherRail) & ",testCondition: " & testCondition & ", it doesn't have any keyword in testCondition to find the matched guardBand for Calculate_Voltage_for_CorePower. Error!!!"
            End If
            
            '''//Check if testCondition contains any keyword about "*GB*".
            If str_testCondition Like "*+*gb*" Then
                voltage_Temp = voltage_Temp + dbl_GB_BinCutJob
            ElseIf str_testCondition Like "*-*gb*" Then 'ex: "MC601 Product-CP2GB"
                voltage_Temp = voltage_Temp - dbl_GB_BinCutJob
            End If
        End If
    End If
        
    '''//Modified for calling the new funciton Calculate_Extra_Voltage_for_PowerRail
    Calculate_Extra_Voltage_for_PowerRail testCondition, voltage_Temp
    
    '''//Output the voltage. It can floor/ceiling the voltage in the functions: "OtherRail_Pwr_Setting", "HVCC_OtherRail_Pwr_Setting".
    voltage_OtherRail = voltage_Temp
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Calculate_Voltage_for_OtherRail"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200508: Modified to merge "Find_Spec_Sheet" and "Find_JobList_Sheet" into "Find_Sheet".
'20200219: Modified to use Lcase(sheetName).
Public Function Find_Sheet(sheetName As String) As Boolean
    Dim strAry_sheetnames() As String
    Dim idxSheet As Integer
    Dim sheettypeSelected As Variant
On Error GoTo errHandler
    '''init
    Find_Sheet = False
    
    '''//Please fined the definition of sheetType in "DataTool.DMGRSheetTypes".
    If LCase(sheetName) Like "*joblist*" Then
        sheettypeSelected = DMGR_SHEET_TYPE_JOBLISTSHEET
    ElseIf LCase(sheetName) Like "dc*specs*" Then
        sheettypeSelected = DMGR_SHEET_TYPE_DCSPECSHEET
    Else
        sheettypeSelected = DMGR_SHEET_TYPE_USER
    End If
    
    strAry_sheetnames = TheExec.Job.GetSheetNamesOfType(sheettypeSelected)
    
    For idxSheet = 0 To UBound(strAry_sheetnames)
        If LCase(strAry_sheetnames(idxSheet)) Like LCase(sheetName) Then
            Find_Sheet = True
        End If
    Next idxSheet
    
    If Find_Sheet = False Then
        TheExec.Datalog.WriteComment sheetName & " doesn't exist in current program which was specified in VBT, Please check the input for Find_Sheet. Error!!!"
        TheExec.ErrorLogMessage sheetName & " doesn't exist in current program which was specified in VBT, Please check the input for Find_Sheet. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Find_Sheet"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Find_Sheet"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210305: Modified CMEM vbt functions to support IDS distribution mode.
'20201218: Modified to remove "enable_CMEM_collection As Boolean" from the arguments of the vbt function "check_flag_to_enable_CMEM_collection".
'20201218: Modified to add "count_FuncPat_decomposed As Long" to the arguments of the vbt function "check_flag_to_enable_CMEM_collection".
'20201218: Modified to move "resize_CMEM_Data_by_pattern_number" from "decide_bincut_feature_for_stepsearch" to "check_flag_to_enable_CMEM_collection".
'20201216: Modified to remove the redundant argument "PrintSize As Long" from the vbt function check_flag_to_enable_CMEM_collection.
'20201215: Modified to check the EnableWord("Vddbin_PTE_Debug")=False to decide "inst_info.enable_CMEM_Collection", suggested by Chihome.
'20201201: Modified to update CaptureSize, failpins, and PrintSize for CMEM.
'20201126: As suggestion from Chihome, modified to set TheHdw.Digital.CMEM.CentralFields for initializing CMEM in GradeSearch_XXX_VT.
'20201111: Created to check the flag "Flag_Enable_CMEM_Collection" to enable CMEM collection.
'20201021: Modified to check DecomposedPat to decide Enable_CMEM_Collection on/off for PatternBurst.
Public Function check_flag_to_enable_CMEM_collection(inst_info As Instance_Info, flag_set_CMEM As Boolean, count_FuncPat_decomposed As Long, CaptureSize As Long, failpins As String)
On Error GoTo errHandler
    '''20201215: Modified to check the EnableWord("Vddbin_PTE_Debug")=False to decide "inst_info.enable_CMEM_Collection", suggested by Chihome.
    '''20210305: Modified CMEM vbt functions to support IDS distribution mode.
    If inst_info.is_BinSearch = True And EnableWord_Vddbin_PTE_Debug = False Then
        If Flag_VDD_Binning_Offline = True Or EnableWord_Vddbinning_OpenSocket = True Then '''If the tester is offline or opensocket.
            inst_info.enable_CMEM_collection = False
        Else
            If flag_set_CMEM = True And count_FuncPat_decomposed > 0 Then
                If inst_info.enable_DecomposePatt = True Then '''Decompose pattern
                    inst_info.enable_CMEM_collection = True
                    
                    '''******************************************************************************************'''
                    '''tlCMEMPatternName    : The pattern name (VM only).
                    '''tlCMEMVMVectorOffset : The vector number (VM only). Store this value only with tlCMEMPatternName.
                    '''tlCMEMModCycle       : The cycle count for the pattern module (VM only).
                    '''******************************************************************************************'''
                    TheHdw.Digital.CMEM.CentralFields = tlCMEMPatternName + tlCMEMVMVectorOffset + tlCMEMModCycle
                Else '''Pattern Burst
                    inst_info.enable_CMEM_collection = False
                End If
                
                '''//If inst_info.enable_CMEM_collection = True, update CaptureSize, failpins, and PrintSize for CMEM for F.C.C. Default value setup.
                If inst_info.enable_CMEM_collection = True Then
                    If failpins = "" Then failpins = "All_Digital"
                    inst_info.PrintSize = CaptureSize
                    If CaptureSize <= 0 Then CaptureSize = 128: inst_info.PrintSize = 1
                End If
            Else '''If FuncPat patset doesn't contain any pattern, disable CMEM_collection
                inst_info.enable_CMEM_collection = False
            End If
        End If
    Else
        inst_info.enable_CMEM_collection = False
    End If

    '''//Only BinSearch uses CMEM to record the failed pattern.
    '''//Clear capture Memory(CMEM) and resize arry of CMEM_data if inst_info.enable_CMEM_Collection is enabled.
    Call resize_CMEM_Data_by_pattern_number(inst_info.enable_CMEM_collection, count_FuncPat_decomposed, inst_info.BC_CMEM_StoreData)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of check_flag_to_enable_CMEM_collection"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of check_flag_to_enable_CMEM_collection"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201202: Modified to add the argument "enable_CMEM_Collection as Boolean" for resize_CMEM_Data_by_pattern_number.
'20201201: Created to resize array of CMEM_data.
'20201126: As suggestion from Chihome, modified 2-dimensions array "Step_CMEM_Data()" and "BC_CMEM_StoreData()" into 1-dimension array to save memory.
Public Function resize_CMEM_Data_by_pattern_number(enable_CMEM_collection As Boolean, count_patternSet_decomposed As Long, current_CMEM_StoreData() As CMEM_StoreData)
    Dim temp_CMEM_Data() As CMEM_StoreData
On Error GoTo errHandler
    If enable_CMEM_collection = True Then
        '''//Initialize CMEM with CmemCaptNone.
        TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone '''CmemCaptNone: Capture no cycles
        
        '''//Resize array of CMEM_StoreData by number of the pattern array.
        If count_patternSet_decomposed > 0 Then
            ReDim temp_CMEM_Data(count_patternSet_decomposed - 1) As CMEM_StoreData
        Else
            ReDim temp_CMEM_Data(0) As CMEM_StoreData
        End If
        
        '''//Resize array of current_CMEM_StoreData.
        current_CMEM_StoreData = temp_CMEM_Data
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of resize_CMEM_Data_by_pattern_number"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of resize_CMEM_Data_by_pattern_number"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190611: Modified for FFC collection by Oscar.
Public Function DoubleToBinaryString(dbl_Input As Double) As String
    Dim bitStream As String
    Dim k As Integer
    Dim dbl_theData As Double ' Used to compare.
    Dim dbl_Temp As Double ' Added I
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''The results stored here as a string that represents which pins have failed.
'''//==================================================================================================================================================================================//'''
    bitStream = vbNullString
    dbl_Temp = dbl_Input
    
    '''Loop and keep shifting the data down.
    For k = 0 To 31
        dbl_theData = Floor(dbl_Temp / 2)
        '''Data is even if theData * 2.0 is equal to Data.
        If dbl_theData * 2 = dbl_Temp Then
            bitStream = "0" + bitStream
        Else
            bitStream = "1" + bitStream
        End If
        dbl_Temp = Floor(dbl_Temp / 2) ' Shift down to next bit.
    Next k
    DoubleToBinaryString = bitStream '''Modified, 20190724
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of DoubleToBinaryString"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190611: Modified for FFC collection by Oscar.
Public Function FailingPins(pinMask As Double, Pins() As String, Offset As Long, PinDict As Dictionary) _
    As String
    Dim i As Long
    Dim strMask As String
    Dim strTemp As String
On Error GoTo errHandler
    If pinMask = 0 Then Exit Function
    
    strMask = DoubleToBinaryString(pinMask)
    FailingPins = vbNullString
    For i = 0 To 31
        strTemp = Mid$(strMask, 32 - i, 1)
        If (i + Offset * 32) > UBound(Pins()) Then Exit For '''Modified, 20190724
        If (strTemp = "1") And PinDict.Exists(Pins(i + Offset * 32)) = False Then
            FailingPins = FailingPins + Pins(i + Offset * 32) + "/"
            PinDict.Add Pins(i + Offset * 32), ""
        End If
    Next i

    '''Get rid of trailing "/"
    If Right(FailingPins, 1) Like "/" Then
        FailingPins = Left$(FailingPins, Len(FailingPins) - 1)
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of FailingPins"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210323: Modified to update Step_CMEM_Data for StoreCapFailcycle if siteResult(site)=False.
'20210317: Modified to use siteResult to decide if Step_CMEM_Data has to be updated.
'20210305: Modified to add the argument "siteResult" to the vbt function "StoreCapFailcycle".
'20210225: Modified to collect CMEM_VectorData/CMEM_CycleData/CMEM_PatRange/CMEM_PatName for each site. Revised by Oscar and PCLINZG.
'20210223: Modified to use site-loop to update Step_CMEM_Data(indexPattern).CMEM_IndexData and Step_CMEM_Data(indexPattern).CMEM_PinData for CMEM.
'20201201: Modified to updated IfStoreData.
'20201126: As suggestion from Chihome, modified 2-dimensions array "Step_CMEM_Data()" and "BC_CMEM_StoreData()" into 1-dimension array to save memory.
'20190611: Added for FFC collection by Oscar.
Public Function StoreCapFailcycle(siteResult As SiteBoolean, failpins As String, indexPattern As Long, CaptureSize As Long, Step_CMEM_Data() As CMEM_StoreData)
    Dim anySiteFailed As Boolean
    '''For temporary CMEM collection
    Dim temp_CMEM_Data As CMEM_StoreData
    Dim temp_CMEM_VectorData() As Double
    Dim temp_CMEM_CycleData() As Double
    Dim temp_CMEM_IndexData As New SiteVariant
    Dim temp_CMEM_PinData As New SiteVariant
    Dim temp_CMEM_PatRange() As Long
    Dim temp_CMEM_PatName() As String
On Error GoTo errHandler
    '''init
    anySiteFailed = False
    
    '''//CapCycle cost 7 ms to retreive even when it is 0
    'Dim CapCycle As Long
    'CapCycle = thehdw.Digital.CMEM.CapturedCycles
    'If CapCycle = 0 Then Exit Function
    
    '''//Get siteResult of pattern pass/fail from the argument "patternPass As SiteBoolean".
    '''//If any site fails on the pattern, set anySiteFailed=True.
    '''20210305: Modified to add the argument "siteResult" to the vbt function "StoreCapFailcycle".
    For Each site In TheExec.sites
        If siteResult(site) = False Then
            anySiteFailed = True
        End If
    Next site
    
    '''//If anySiteFailed = True, collect data from CMEM.
    If anySiteFailed = True Then
        '''//This method gets the per-site fail cycle scan data for the specified pins.
        '''========================================================================================================================================================================================'''
        '''IndexData: ISiteVariant. A site variable object that contains array of index values that represents the scan data fail locations. IG-XL uses this object to index into the central CMEM.
        '''PinData  : ISiteVariant. A site variable object that contains an array of Doubles that represent which pins have failed for a given cycle.
        '''========================================================================================================================================================================================'''
        '''ToDo: Maximum Cycles Captured of UltraPin1600 is 256K major cycles. Maybe we can use globalVariable to define Capture Cycle to save memory...
        Call TheHdw.Digital.Pins(failpins).CMEM.StoredCycleData(temp_CMEM_Data.CMEM_IndexData, temp_CMEM_Data.CMEM_PinData, -1, False)
    
        '''//Read back the last index where data was captured for each pattern.
        Call TheHdw.Digital.CMEM.PatternName(temp_CMEM_PatRange, temp_CMEM_PatName)
        
        '''//This property gets the pattern generator information stored in the capture memory (CMEM).
        '''***************************************************************************************************************************'''
        '''tlCMEMVMVectorOffset : The vector number (VM only). Store this value only with tlCMEMPatternName.
        '''tlCMEMModCycle       : The cycle count for the pattern module (VM only).
        '''***************************************************************************************************************************'''
        temp_CMEM_VectorData = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMVMVectorOffset, -1)
        temp_CMEM_CycleData = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMModCycle, -1)
        
        '''//Save StoredCycleData from CMEM for current Pattern into Step_CMEM_Data.
        '''20210225: Modified to collect CMEM_VectorData/CMEM_CycleData/CMEM_PatRange/CMEM_PatName for each site. Revised by Oscar and PCLINZG.
        '''20210323: Modified to update Step_CMEM_Data for StoreCapFailcycle if siteResult(site)=False.
        For Each site In TheExec.sites
            If siteResult(site) = False Then
                Step_CMEM_Data(indexPattern).CMEM_IndexData = temp_CMEM_Data.CMEM_IndexData '''SiteVariant
                Step_CMEM_Data(indexPattern).CMEM_PinData = temp_CMEM_Data.CMEM_PinData '''SiteVariant
                Step_CMEM_Data(indexPattern).CMEM_VectorData = temp_CMEM_VectorData
                Step_CMEM_Data(indexPattern).CMEM_CycleData = temp_CMEM_CycleData
                Step_CMEM_Data(indexPattern).CMEM_PatRange = temp_CMEM_PatRange
                Step_CMEM_Data(indexPattern).CMEM_PatName = temp_CMEM_PatName
                '''//Set IfStoreData=True. The flag "IfStoreData" will be checked in the vbt functions StoreCaptureByStep and PostTestIPF.
                Step_CMEM_Data(indexPattern).IfStoreData(site) = True
            End If
        Next site
    Else '''If no site fails on the pattern, exit the function.
        Exit Function
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of StoreCapFailcycle"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210305: Modified to add the arguments "step_control As Instance_Step_Control" to the vbt function "StoreCaptureByStep".
'20210225: Modified to collect CMEM_VectorData/CMEM_CycleData/CMEM_PatRange/CMEM_PatName for each site. Revised by Oscar and PCLINZG.
'20210223: Modified to use site-loop to update BC_CMEM_Data(indexPattern).CMEM_IndexData and BC_CMEM_Data(indexPattern).CMEM_PinData for CMEM.
'20201201: Modified to add the argument "BC_CMEM_StoreData() As CMEM_StoreData".
'20201126: As suggestion from Chihome, modified 2-dimensions array "Step_CMEM_Data()" and "BC_CMEM_StoreData()" into 1-dimension array to save memory.
'20190611: Added for FFC collection by Oscar.
Public Function StoreCaptureByStep(inst_info As Instance_Info, Step_CMEM_Data() As CMEM_StoreData, BC_CMEM_StoreData() As CMEM_StoreData)
    Dim site As Variant
    Dim Clear_CMEM_Data As CMEM_StoreData
    Dim indexPattern As Long
    Dim anySiteFailed As Boolean
On Error GoTo errHandler
    '''init
    anySiteFailed = False

    For Each site In TheExec.sites
        If inst_info.All_Patt_Pass(site) = False Then
            anySiteFailed = True
        End If
    Next site
    
    If anySiteFailed = True Then
        For indexPattern = 0 To UBound(Step_CMEM_Data)
            '''//Copy CMEM data from Step_CMEM_Data to BC_CMEM_StoreData for each pattern.
            '''20210225: Modified to collect CMEM_VectorData/CMEM_CycleData/CMEM_PatRange/CMEM_PatName for each site. Revised by Oscar and PCLINZG.
            '''20210305: Modified to revise the vbt function StoreCaptureByStep for IDS distribution mode. CMEM should be updated while grade_found = False.
            For Each site In TheExec.sites
                If Step_CMEM_Data(indexPattern).IfStoreData(site) = True _
                And ((inst_info.gradeAlg(site) = GradeSearchAlgorithm.linear And inst_info.All_Patt_Pass(site) = False) Or (inst_info.gradeAlg(site) = GradeSearchAlgorithm.IDS And inst_info.All_Patt_Pass(site) = False And inst_info.grade_found(site) = False)) Then
                    BC_CMEM_StoreData(indexPattern).CMEM_IndexData = Step_CMEM_Data(indexPattern).CMEM_IndexData '''SiteVariant
                    BC_CMEM_StoreData(indexPattern).CMEM_PinData = Step_CMEM_Data(indexPattern).CMEM_PinData '''SiteVariant
                    BC_CMEM_StoreData(indexPattern).IfStoreData = Step_CMEM_Data(indexPattern).IfStoreData '''SiteVariant
                    BC_CMEM_StoreData(indexPattern).CMEM_VectorData = Step_CMEM_Data(indexPattern).CMEM_VectorData
                    BC_CMEM_StoreData(indexPattern).CMEM_CycleData = Step_CMEM_Data(indexPattern).CMEM_CycleData
                    BC_CMEM_StoreData(indexPattern).CMEM_PatRange = Step_CMEM_Data(indexPattern).CMEM_PatRange
                    BC_CMEM_StoreData(indexPattern).CMEM_PatName = Step_CMEM_Data(indexPattern).CMEM_PatName
                End If
            Next site
        Next indexPattern
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of StoreCaptureByStep"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210225: Modified to collect CMEM_VectorData/CMEM_CycleData/CMEM_PatRange/CMEM_PatName for each site. Revised by Oscar and PCLINZG.
'20210125: Modified to revise the message about "Exceeds DTR limitation", requested by ZQLIN.
'20201126: As suggestion from Chihome, modified 2-dimensions array "Step_CMEM_Data()" and "BC_CMEM_StoreData()" into 1-dimension array to save memory.
'20190611: Added for FFC collection by Oscar.
Public Function PostTestIPF(performance_mode As String, failpins As String, CaptureSize As Long, Temp_BC_CMEM_StoreData() As CMEM_StoreData)
    '''Decompose the pins to be used in reading out the failures.
    Dim cnt_pins As Long
    Dim Pins() As String
    Dim inst_name As String
    '''Variables used in site loop
    Dim site As Variant
    Dim indexPattern As Long
    Dim patternNames() As String
    Dim lastFailPerPattern() As Long
    Dim indexFailedPattern As Long
    Dim PinStr As String
    Dim i As Long
    Dim k As Long
    Dim M As Long
    Dim bigPinList As Integer ' Used when pinlist is greater than 32.
    Dim indexArr() As Long ' Used in site loop
    Dim pinDataArr() As Double ' Used in site loop
    Dim tempPinlist As String
    Dim SiteString As String
    Dim PatternString As String
    Dim headString As String
    Dim printstring As String
    Dim titleString As String
    Dim StringLength As Long
    Dim IfPatChange As Boolean
    Dim PreVector As Double
    Dim IfVectorChange As Boolean
    Dim PinDict As New Dictionary
    Dim temp_pinarr() As String
    Dim vectorString As String
    Dim cycleString As String
    Dim ExtendPin As Boolean
On Error GoTo errHandler
    '''//Get the instance name.
    inst_name = TheExec.DataManager.instanceName

    '''//Decompose pinGroup from the argument "failpins".
    Call TheExec.DataManager.DecomposePinList(failpins, Pins, cnt_pins)

    For indexPattern = 0 To UBound(Temp_BC_CMEM_StoreData)
        For Each site In TheExec.sites
            If Temp_BC_CMEM_StoreData(indexPattern).IfStoreData = False Then GoTo nextSite
            
            patternNames = Temp_BC_CMEM_StoreData(indexPattern).CMEM_PatName
            lastFailPerPattern = Temp_BC_CMEM_StoreData(indexPattern).CMEM_PatRange
            PreVector = -1

            '''//Move siteVariant to normal array of Long/Double from the array "Temp_BC_CMEM_StoreData".
            '''========================================================================================================================================================================================'''
            '''IndexData: ISiteVariant. A site variable object that contains array of index values that represents the scan data fail locations. IG-XL uses this object to index into the central CMEM.
            '''PinData  : ISiteVariant. A site variable object that contains an array of Doubles that represent which pins have failed for a given cycle.
            '''========================================================================================================================================================================================'''
            pinDataArr = Temp_BC_CMEM_StoreData(indexPattern).CMEM_PinData '''siteVariant
            indexArr = Temp_BC_CMEM_StoreData(indexPattern).CMEM_IndexData '''siteVariant

            '''//The result may be pass, so that CMEM data would be empty.
            If UBound(indexArr) = -1 Then GoTo nextSite

            headString = 1
            SiteString = CStr(site)
            titleString = "FFC1" + "," + headString + "," + SiteString + "," + "BV_" + performance_mode + "," + inst_name
            printstring = titleString

            '''//bigPinlist will equal 1 for <32 pins; 2 for 64 to 33 pins; and so on.
            bigPinList = (UBound(pinDataArr) + 1) \ (UBound(indexArr) + 1)

            indexFailedPattern = 0
            StringLength = 0
            ExtendPin = False
            'thehdw.StartStopwatch

            For i = LBound(indexArr) To UBound(indexArr)
                PinStr = vbNullString
                If indexArr(i) > UBound(Temp_BC_CMEM_StoreData(indexPattern).CMEM_VectorData(site)) Then
                    TheExec.Datalog.WriteComment "site:" & site & "," & "Capture Data exceeds the maximum vector which was stored by CMEM"
                    Exit For
                End If

                If indexFailedPattern > UBound(patternNames) Then
                    '''Raise error. This should never happen.
                End If

                If i > lastFailPerPattern(indexFailedPattern) Then
                    '''//Go to the next pattern in the list.
                    indexFailedPattern = indexFailedPattern + 1
                    IfPatChange = True
                End If

                '''//patternname vectoroffset scanindex failingpins
                PatternString = patternNames(indexFailedPattern)
                vectorString = CStr(Temp_BC_CMEM_StoreData(indexPattern).CMEM_VectorData(site)(indexArr(i)))
                cycleString = CStr(Temp_BC_CMEM_StoreData(indexPattern).CMEM_CycleData(site)(indexArr(i)))

                If Not (PreVector Like vectorString) And IfPatChange = False Then
                    PreVector = vectorString
                    IfVectorChange = True
                    PinDict.RemoveAll
                End If

                For k = 0 To bigPinList - 1
                    tempPinlist = FailingPins(pinDataArr(i * bigPinList + k), Pins, k, PinDict)
                    If tempPinlist <> "" Then PinStr = tempPinlist & "/" & PinStr
                Next k

                '''//Get rid of trailing "/"
                If Right(PinStr, 1) Like "/" Then
                    PinStr = Left$(PinStr, Len(PinStr) - 1)
                ElseIf PinStr = "" Then
                    GoTo nexti
                End If

                Dim TempPin As Variant
                Dim TempPinArr() As String
                Dim PinSplit As String

                If ((Len(titleString + PatternString + vectorString + cycleString + PinStr) + 4) > 250) Then
                    '''//If this line would be over 250 strings by itself even when we switch to next line. Initializing "Extend Pin" Mode.
                    If Len(titleString + PatternString) > 250 Then TheExec.Datalog.WriteComment "Error, Title & Pattern exceed DTR limitation."

                    '''//If this line were desitined to be over 250 strings. Print Error message , and still, we use "Extend Pin" Mode to print as far as we can.
                    '''20210125: Modified to revise the message about "Exceeds DTR limitation", requested by ZQLIN.
                    TheExec.Datalog.WriteComment "Current Cycle: (" + vectorString + ", " + cycleString + ")'s Fail info. Exceeds DTR limitation, Starting to use Decompose Pin Mode. Warning!!!"

                    TempPinArr = Split(PinStr, "/")

                    For Each TempPin In TempPinArr
                        If (printstring = titleString) Then
                            printstring = titleString + "," + PatternString + "," + vectorString + "," + cycleString + "," + TempPin
                        Else
                            printstring = printstring + vbCrLf + titleString + "," + PatternString + "," + vectorString + "," + cycleString + "," + TempPin
                        End If
                    Next TempPin

                    ExtendPin = True
                    '''========================================================================'''
                    '''Rule: Next Cycle would not be allowed to append directly.
                    '''StringLength = Len(titleString)
                    '''No Need to Initialize the String Length since we don't allow it to append.
                    '''========================================================================'''
                ElseIf (ExtendPin = True) _
                Or ((StringLength + Len(vectorString + cycleString + PinStr) + 3) > 250) _
                Or ((StringLength + Len(PinStr) + 1) > 250) _
                Or (IfPatChange = True) Then
                    '''========================================================================'''
                    '''If Previous Cycle were "Extend Pin" Mode.
                    '''If Directly Data-Appending(Different Vector/Same Vector) would be over 250 strings.
                    '''If it were Different Pattern.
                    '''========================================================================'''
                    printstring = printstring + vbCrLf + titleString + "," + PatternString + "," + vectorString + "," + cycleString + "," + PinStr

                    '''//We already make sure it would be less than 250 by itself, so we can just use new line to print.
                    StringLength = Len(titleString + PatternString + vectorString + cycleString + PinStr) + 4

                    '''//Succeed from this new line for next one to append the data.
                    ExtendPin = False

                ElseIf (IfVectorChange = True) Then
                    If (printstring = titleString) Then printstring = printstring + "," + PatternString: StringLength = Len(titleString + PatternString) + 1

                    '''//First Line would be considered as new vector without Pattern String since initial string is only "TitleString".
                    printstring = printstring + "," + vectorString + "," + cycleString + "," + PinStr
                    StringLength = StringLength + Len(vectorString + cycleString + PinStr) + 3
                    ExtendPin = False
                Else
                    printstring = printstring + "/" + PinStr
                    '''//If it is the same vector at same pattern, and it would be less than 250 if we append the PinStr
                    StringLength = StringLength + Len(PinSplit) + 1
                    ExtendPin = False
                End If

                IfPatChange = False
                IfVectorChange = False

                '''//Ensure we don't get "suscript out of index" error.
                'If i >= CaptureSize Then Exit For
                If i >= CaptureSize - 1 Then Exit For
nexti:
            Next i
            'Debug.Print "Print:" & thehdw.ReadStopwatch
            If printstring <> titleString Then TheExec.Datalog.WriteComment printstring
nextSite:
        Next site
nextindexPattern:
    Next indexPattern
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of PostTestIPF"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200703: Modiifed to use "check_Sheet_Range".
'20200528: Modified to check header of the table.
'20200508: Modified to merge "Find_Spec_Sheet" and "Find_JobList_Sheet" into "Find_Sheet".
'20191231: Created to parse the sheet "GeneratedJobListSheet".
Public Function Parsing_GeneratedJobList_Sheet(testJob As String, JobListSheet As String, DcSpecsSheet As String)
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim i As Integer, j As Integer
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim row_of_title As Long
    Dim col_testJob As Long
    Dim col_DcSpecSheet As Long
    Dim binNumber As Integer
    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
On Error GoTo errHandler
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "GeneratedJobListSheet"
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(JobListSheet, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        enableRowParsing = False
        
        '''//Check the header of the table
        For i = 1 To MaxRow
            For j = 1 To maxcol
                If UCase(ws_def.Cells(i, j)) Like UCase("Job*Name") Then
                    col_testJob = j
                    row_of_title = i
                End If
                
                If row_of_title > 0 Then
                    If UCase(ws_def.Cells(i, j)) Like UCase("DC*Specs") Then
                        col_DcSpecSheet = j
                    End If
                End If
                
                If col_testJob > 0 And col_DcSpecSheet > 0 Then
                    Exit For
                End If
            Next j
            
            If row_of_title > 0 Then
                If col_testJob > 0 And col_DcSpecSheet > 0 Then
                    enableRowParsing = True
                    Exit For
                Else
                    enableRowParsing = False
                    TheExec.Datalog.WriteComment "The header of " & JobListSheet & " is incorrect. Error!!!"
                    TheExec.ErrorLogMessage "The header of " & JobListSheet & " is incorrect. Error!!!"
                End If
            End If
        Next i
        
        '''//Parsing all powerBinning sheets.
        If enableRowParsing = True Then
            For i = row_of_title + 1 To MaxRow
                If UCase(ws_def.Cells(i, col_testJob).Value) Like UCase(testJob) Then '''According to testJob, find the matched DC Specs sheet.
                    If UCase(ws_def.Cells(i, col_DcSpecSheet).Value) <> "" Then
                        DcSpecsSheet = ws_def.Cells(i, col_DcSpecSheet).Value
                        Exit For
                    Else
                        DcSpecsSheet = ""
                        TheExec.Datalog.WriteComment "No matched DC Specs in the sheet " & JobListSheet & " for testJob " & testJob & ". Error!!!"
                        TheExec.ErrorLogMessage "No matched DC Specs in the sheet " & JobListSheet & " for testJob " & testJob & ". Error!!!"
                    End If
                End If
            Next i
        Else
            TheExec.Datalog.WriteComment sheetName & " doesn't have the correct header for Parsing_GeneratedJobList_Sheet. Error!!!"
            TheExec.ErrorLogMessage sheetName & " doesn't have the correct header for Parsing_GeneratedJobList_Sheet. Error!!!"
        End If
    End If '''If isSheetFound = True
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Parsing_GeneratedJobList_Sheet"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Parsing_GeneratedJobList_Sheet"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210325: Modified to use the 1-dimension array to store SRAM_Vth.
'20201116: Modified to remove the redunant "Ucase()".
'20201005: Modified to check if powerDomain is "PRESERVED" or "RESERVED". C651 Toby requested this on 20200916.
'20200106: Created to check SRAMthresh for p_mode of selsram powerpin.
'20180821: Modified to check bincut testjob mapping.
Public Function Precheck_SRAMthresh_for_Selsram_Power(powerGroup As String)
    Dim sram_Vth As Double
    Dim p_mode As Integer
    Dim split_powerGroup() As String
    Dim powerDomain As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim bin_num As Long
    Dim idx_sramVth As Long
On Error GoTo errHandler
    If powerGroup <> "" Then
        '''//Check each powerDomain in powerGroup
        split_powerGroup = Split(powerGroup, ",")
        
        For i = 0 To UBound(split_powerGroup)
            powerDomain = UCase(Trim(split_powerGroup(i)))
            
            '''//Check if powerDomain is "PRESERVED" or "RESERVED". C651 Toby requested this on 20200916.
            If powerDomain <> "PRESERVED" And powerDomain <> "RESERVED" Then
                '''//Check each p_mode of powerDomain.
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq)
                    p_mode = VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq(j))
                    
                    '''//Check SRAM_VTH_SPEC for p_mode by each bin_num.
                    For k = 0 To UBound(PassBinCut_ary)
                        bin_num = PassBinCut_ary(k)
                        '''20210325: Modified to use the 1-dimension array to store SRAM_Vth.
                        '''************************************************************************************************************************'''
                        '''column "SRAMthresh_CP1" or "SRAMthresh_BinSearch"==> SRAM_VTH_SPEC(0): for CP1 BV binSearch and postBinCut/OutsideBinCut.
                        '''column "SRAMthresh_Product"                      ==> SRAM_VTH_SPEC(1): for CP1 HBV and non-CP1 BV/HBV.
                        '''************************************************************************************************************************'''
                        For idx_sramVth = 0 To UBound(BinCut(p_mode, bin_num).SRAM_VTH_SPEC)
                            sram_Vth = BinCut(p_mode, bin_num).SRAM_VTH_SPEC(idx_sramVth)
                        
                            If sram_Vth > 0 Then
                                '''Do nothing
                            Else
                                TheExec.Datalog.WriteComment VddBinName(p_mode) & " of powerDomain " & powerDomain & " has incorrect SRAMthresh in Bin" & bin_num & ". Please check Vdd_Binning_Def_appA" & bin_num & ". Error!!!"
                                TheExec.ErrorLogMessage VddBinName(p_mode) & " of powerDomain " & powerDomain & " has incorrect SRAMthresh in Bin" & bin_num & ". Please check Vdd_Binning_Def_appA" & bin_num & ". Error!!!"
                            End If
                        Next idx_sramVth
                    Next k
                Next j
            End If
        Next i
    Else
        TheExec.Datalog.WriteComment "The argument of powerGroup of Precheck_SRAMthresh_for_Selsram_Power should not be empty. Please check SELSRM_Mapping_Table. Error!!!"
        TheExec.ErrorLogMessage "The argument of powerGroup of Precheck_SRAMthresh_for_Selsram_Power should not be empty. Please check SELSRM_Mapping_Table. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Precheck_SRAMthresh_for_Selsram_Power"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Precheck_SRAMthresh_for_Selsram_Power"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210729: Modified to remove the redundant argument "sram_Vth() As SiteDouble" from the vbt functions Get_BitSequence_for_CorePower and Write_Selsrm_DSSC_Info.
'20210728: Modified to move "Dim Sram_Vth(MaxBincutPowerdomainCount) As New SiteDouble" into "Public Type Instance_Info".
'20201209: Modified to remove the argument "ByRef voltage_SelsrmBitCalc() As SiteDouble" and use "inst_info.voltage_SelsrmBitCalc" for Calculate_Selsrm_DSSC_For_BinCut.
'20201029: Modified to use "Public Type Instance_Info".
'20201012: Modified to replace "DSSC_init_pat" with "selSramPat".
'20201012: Modified to replace "pin_write" with "selsrm_DigSrc_Pin".
'20200824: Modified to use Expand_SelsramBit_to_VddbinDsscArray and Write_Selsrm_DSSC_Info.
'20200821: Modified to use the siteLong array for bit calculation of CorePower.
'20200821: Modified to check SrmDSSC and DSrmDSSC.
'20200802: Modified to use "pinmappingDic" to store SELSRM bits of corePower.
'20200728: Modified to remove the redundant pin-loop for init the sideDouble array SRAM_Vth".
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20200212: Modified to clear the siteVariable "SelsramInfo" before SELSRM DSSC calculation.
'20200203: Modified to calculate SELSRM DSSC bits by comparing the applied Payload voltages (real voltages) with SRAM_Vth (suggestion from C651 Chris Vu).
'20200130: Modified to move SELSRM DSSC bits calculation from Non_Binning_Pwr_Setting_VT / HVCC_Set_VT / PostBinCut_Voltage_Set_VT to Calculate_Selsrm_DSSC_Bits_For_BinCut.
'20200107: Modified to use "Expand_SelsramBit_to_VddbinDsscArray".
'20191119: Modified to use pattern names to find DSSC bit sequence from SELSRM_Mapping_Table.
Public Function Calculate_Selsrm_DSSC_For_BinCut(inst_info As Instance_Info, passBinCut As SiteLong)
    Dim site As Variant
    Dim i As Long
    Dim bitSeq_CorePower(MaxSiteCount - 1) As String
    Dim bitArray_corePower(MaxBincutPowerdomainCount) As New SiteLong
    Dim powerDomain As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Please check bit numbers of SelSram (SELSRAM_Mapping_Table) before you modify this block!!!
'''2. voltage_SelsrmBitCalc is the siteDouble array to store voltage for Selsrm bit calculation.
'''3. str_Selsrm_DSSC_Info is used to store string about SELSRM bits comparison.
'''4. str_Selsrm_DSSC_Bit is used to store string about SELSRM bits sequence(LSB->MSB) of each site.
'''
'''//voltage_SelsrmBitCalc
'''20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'''CP1 BV: CP1 Grade Search     : Compare BV to SRAMthresh_binSearch (formerly SRAMthresh_CP1). => DO NOT include Offsets or voltage delta (+15mV) in the comparison.
'''CP1 HBV and non-CP1 BV/HBV   : After CP1 Grade Search: Compare PV fused voltage to SRAMthresh_Product. This is for CP1 HBV, CP2, WLFT1, FT1/2/3/, etc... => DO NOT include Offsets, or voltage delta (+15mV) in the comparison.
'''//==================================================================================================================================================================================//'''
    '''//Init
    VBIN_RESULT(inst_info.p_mode).DSSC_Dec = -1
    
    '''//Initialize the siteLong array of CorePower for SELSRM DSSC.
    For i = 0 To UBound(bitArray_corePower)
        bitArray_corePower(i) = 0
    Next i
    
    '''//Check if selsrm_DigSrc_SignalName and selsrm_DigSrc_Pin are provided.
    If inst_info.selsrm_DigSrc_SignalName = "" Then
        TheExec.Datalog.WriteComment "inst_info.selsrm_DigSrc_SignalName of Expand_SelsramBit_to_VddbinDsscArray for " & inst_info.inst_name & " is empty. Error!!!"
    End If
    
    If inst_info.selsrm_DigSrc_Pin = "" Then
        TheExec.Datalog.WriteComment "inst_info.selsrm_DigSrc_Pin of Expand_SelsramBit_to_VddbinDsscArray for " & inst_info.inst_name & " is empty. Error!!!"
    End If

    '''//Init the strings of Sram_Vth and siteVariant for each site.
    For Each site In TheExec.sites
        inst_info.str_Selsrm_DSSC_Info(site) = ""
        bitSeq_CorePower(site) = ""
    Next site
    
    '''***********************************************************************************'''
    '''selSramPat is provided by GradeSearch_xxx_VT.
    '''If init_pat is not empty, it means that DUT uses SelSram and VRS.
    '''If init_pat is empty, it means that DUT uses the conventional method, no DSSC SelSram.
    '''***********************************************************************************'''
    If inst_info.patt_SelsrmDigSrc_single <> "" Then
        '''//If the pattern has the matched pattern keyword of block, start voltage comparison for bit calculation.
        If inst_info.idxBlock_Selsrm_singlePatt > -1 Then
            '''============================================================================================================================================
            '''[step1] Get bit sequence for BinCut CorePower by comparing Selsrm Logic Power with Sram_Vth (from the sheet "Vdd_Binning_Def").
            '''============================================================================================================================================
            Call Get_BitSequence_for_CorePower(inst_info, passBinCut, inst_info.voltage_SelsrmBitCalc, bitArray_corePower)

            '''============================================================================================================================================
            '''[step2] Expand bit_CorePower_array into SelSram DSSC DigSrc array.
            '''============================================================================================================================================
            Call Expand_SelsramBit_to_VddbinDsscArray(inst_info, bitArray_corePower)
            
            '''============================================================================================================================================
            '''[step3] Collect information about voltage_SelsrmBitCalc, Sram_Vth, DSSC, SELSRAM_Compare_Bit_Str, and SELSRAM_DSSC_Bit_Str into the string arrary for datalog.
            '''============================================================================================================================================
            Call Write_Selsrm_DSSC_Info(inst_info, inst_info.voltage_SelsrmBitCalc, bitArray_corePower)
        Else
            TheExec.Datalog.WriteComment inst_info.patt_SelsrmDigSrc_single & " doesn't have any idx_block about Selsrm_Mapping_Table for Calculate_Selsrm_DSSC_For_BinCut. Error!!!"
            TheExec.ErrorLogMessage inst_info.patt_SelsrmDigSrc_single & " doesn't have any idx_block about Selsrm_Mapping_Table for Calculate_Selsrm_DSSC_For_BinCut. Error!!!"
        End If
    Else '''If selSramPat <> "" Then
        '''//selSramPat = "", it is the static pattern (No DSSC for SELSRAM DSSC DigSrc).
        '''If Pattern set without any DSSC pattern, skip SELSRM DSSC calculation and set DSSC_Dec=-1.
        VBIN_RESULT(inst_info.p_mode).DSSC_Dec = -1
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Calculate_Selsrm_DSSC_Bits_For_BinCut"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210729: Modified to remove the redundant argument "sram_Vth() As SiteDouble" from the vbt functions Get_BitSequence_for_CorePower and Write_Selsrm_DSSC_Info.
'20210729: Modified to remove the redundant function Get_SRAM_VTH because SRAM_Vth is determined by Calculate_LVCC_for_CorePower or Calculate_HVCC_for_CorePower, as suggested by C651 and Chihome.
'20210728: Modified to move "Dim Sram_Vth(MaxBincutPowerdomainCount) As New SiteDouble" into "Public Type Instance_Info".
'20210720: Modified to update the branch to check if check if CorePower is the selsrm LogicPin.
'20210719: Modified to check if CorePower is the selsrm LogicPin, as requested by ZYLINI.
'20210628: Modified to revise the rules for SRM and DSRM as discussion with TSMC SYLIUU.
'20210129: Modified the branches of "If isDsrmDSSC = True".
'20210126: Modified to revise the vbt code for DevChar.
'20201029: Modified to use "Public Type Instance_Info".
'20200824: Created to get bit sequence for CorePower by comparing payload voltages of BinCut corePower with Sram_Vth.
'20200821: Modified to check SrmDSSC and DSrmDSSC.
'20200728: Modified to remove the redundant pin-loop for init the sideDouble array "SRAM_Vth".
Public Function Get_BitSequence_for_CorePower(inst_info As Instance_Info, passBinCut As SiteLong, voltage_SelsrmBitCalc() As SiteDouble, bitArray_corePower() As SiteLong)
    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim powerDomain As String
    Dim pin_SramPower As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''20210628: Modified to revise the rules for SRM and DSRM as discussion with TSMC SYLIUU.
'''For pattern with keyword "*SRM*" (SRM or DSRM): if (Logic corePower > Sram_Vth) --> get bit from column "Selsrm0" of sheet "SelSrm_Mapping_Table".
'''0: means column selsrm0; 1: means column selsrm1.
'''//==================================================================================================================================================================================//'''
    For i = 0 To UBound(pinGroup_CorePower)
        powerDomain = UCase(pinGroup_CorePower(i))
        
        '''//Check if CorePower is the selsrm LogicPin.
        '''20210719: Modified to check if CorePower is the selsrm LogicPin, as requested by ZYLINI.
        '''20210720: Modified to update the branch to check if check if CorePower is the selsrm LogicPin.
        If LCase("*," & selsramLogicPin & ",*") Like LCase("*," & powerDomain & ",*") Then
            '''//Get Sram_Vth for each p_mode with bin number.
            If inst_info.is_DevChar_Running = True Then
                '''init
                pin_SramPower = ""
                
                For j = 0 To UBound(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).logic_Pin)
                    If UCase(powerDomain) = UCase(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).logic_Pin(j)) Then
                        '''//Use real BinCut Payload voltages of Logic Power and SRAM Power for DevChar.
                        '''Logic Power
                        voltage_SelsrmBitCalc(VddBinStr2Enum(powerDomain)) = BinCut_Payload_Voltage(VddBinStr2Enum(powerDomain))
                        '''SRAM Power
                        pin_SramPower = SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).sram_Pin(j)
                        inst_info.sram_Vth(VddBinStr2Enum(powerDomain)) = BinCut_Payload_Voltage(VddBinStr2Enum(pin_SramPower))
                        Exit For
                    End If
                Next j
            End If
            
            '''//Compare corePower voltages with sram_vth for Bit calculation.
            '''20210628: Modified to revise the rules for SRM and DSRM as discussion with TSMC SYLIUU.
            '''0: means column selsrm0; 1: means column selsrm1.
            For Each site In TheExec.sites
                bitArray_corePower(VddBinStr2Enum(powerDomain)) = IIf(CDec(voltage_SelsrmBitCalc(VddBinStr2Enum(powerDomain))) > CDec(inst_info.sram_Vth(VddBinStr2Enum(powerDomain))), 0, 1)
            Next site
        End If
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Get_Bit_for_CorePower"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Get_Bit_for_CorePower"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210628: Modified to revise the rules for SRM and DSRM as discussion with TSMC SYLIUU.
'20201116: Modified to remove the redunant "Ucase()".
'20201029: Modified to use "Public Type Instance_Info".
'20201012: Modified to replace "DSSC_init_pat" with "selSramPat".
'20201012: Modified to replace "pin_write" with "selsrm_DigSrc_Pin".
'20200922: Modified to check "PRESERVED" and "RESERVED". C651 Toby requested this on 20200916.
'20200824: Created to expand bit array of BinCut payload voltages into Selsrm DSSC DigSrc bit array.
Public Function Expand_SelsramBit_to_VddbinDsscArray(inst_info As Instance_Info, bitArray_corePower() As SiteLong)
    Dim site As Variant
    Dim i As Integer
    Dim powerDomain As String
    Dim selsram_expand_array() As Long
On Error GoTo errHandler
    ReDim selsram_expand_array(UBound(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).logic_Pin))
    
'Dim test_time1 As Long
'test_time1 = Timer
TheHdw.Patterns(inst_info.patt_SelsrmDigSrc_single).Load
    For Each site In TheExec.sites
        For i = 0 To UBound(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).logic_Pin)
            powerDomain = UCase(Trim(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).logic_Pin(i)))
        
            '''//Check if powerDomain is "PRESERVED" or "RESERVED". C651 Toby requested this on 20200916.
            If powerDomain Like "PRESERVED" Or powerDomain Like "RESERVED" Then
                selsram_expand_array(i) = CInt(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).SelSrm0(i))
            Else
                '===========================================================================================================================================
                '''//For pattern with keyword "*SRM*" (SRM or DSRM), if (Logic corePower > Sram_Vth) --> get bit from column "Selsrm0" of sheet "SelSrm_Mapping_Table".
                '''20210628: Modified to revise the rules for SRM and DSRM as discussion with TSMC SYLIUU.
                '''SelSrm0(i): means column selsrm0; SelSrm1(i): means column selsrm1.
                '===========================================================================================================================================
                selsram_expand_array(i) = IIf(CInt(bitArray_corePower(VddBinStr2Enum(powerDomain))), CInt(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).SelSrm1(i)), CInt(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).SelSrm0(i)))
            End If
        Next i
            
       
        '''//Setup DSSC DigSrc with selSramPat and selsrm_DSSC_DigSrc_array.
        DSSC_SetupDigSrcWave inst_info.patt_SelsrmDigSrc_single, inst_info.selsrm_DigSrc_Pin, inst_info.selsrm_DigSrc_SignalName, UBound(selsram_expand_array) + 1, selsram_expand_array
       

        '''//Merge the array into the string of SelSrm DSSC DigSrc.
        inst_info.str_Selsrm_DSSC_Bit(site) = ""
        
        For i = 0 To UBound(selsram_expand_array)
            inst_info.str_Selsrm_DSSC_Bit(site) = inst_info.str_Selsrm_DSSC_Bit(site) & CStr(selsram_expand_array(i))
        Next i

        '''//VBIN_RESULT(p_mode).DSSC_Dec will printed in the datalog when Judge_PF.
        VBIN_RESULT(inst_info.p_mode).DSSC_Dec = auto_bitStr2Dec(inst_info.str_Selsrm_DSSC_Bit(site), True)
    Next site
    
'TheExec.Datalog.WriteComment ("***** Test Time (VBA): DSSC_SetupDigSrcWave (s) = " & Format(Timer - test_time1, "0.000000"))
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Expand_SelsramBit_to_VddbinDsscArray"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Expand_SelsramBit_to_VddbinDsscArray"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210729: Modified to remove the redundant argument "sram_Vth() As SiteDouble" from the vbt functions Get_BitSequence_for_CorePower and Write_Selsrm_DSSC_Info.
'20210219: Modified to check the flag "Flag_Skip_Printing_SelSrm_DSSC_Info" to skip printing BV SelSrm Info.
'20201116: Modified to remove the redunant "Ucase()".
'20201029: Modified to use "Public Type Instance_Info".
'20200922: Modified to check "PRESERVED" and "RESERVED". C651 Toby requested this on 20200916.
'20200831: Modified to use "Flag_Remove_Printing_BV_voltages".
'20200831: Modified to skip "PRESERVED".
'20200824: Created to write strings of Selsrm DSSC DigSrc into datalog.
Public Function Write_Selsrm_DSSC_Info(inst_info As Instance_Info, voltage_SelsrmBitCalc() As SiteDouble, bitArray_corePower() As SiteLong)
    Dim site As Variant
    Dim i As Integer
    Dim powerDomain As String
    Dim str_bitCompareAll As String
    Dim str_bitCompare As String
    'Dim str_bitDSSC As String
    Dim str_sram_vth As String
    Dim str_logic_voltage As String
    Dim alphaPowerDomain As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Example:
'''SELSRAM_Compare_Bit_Str,0,11111(LSB->MSB),S=1,G=1,I=1,D=1,A=1
'''SELSRAM_DSSC_Bit_Str,0,00000
'''SRAM_Vth(DCVS),0,S=0.780V,G=0.765V,I=0.775V,D=0.765V,A=0.770V
'''SelSram_voltage,0,S=0.550V,G=0.630V,I=0.615V,D=0.615V,A=0.615V
'''//==================================================================================================================================================================================//'''
    If Flag_Remove_Printing_BV_voltages = False And Flag_Skip_Printing_SelSrm_DSSC_Info = False Then
        For Each site In TheExec.sites
            '''//Prefix
            str_bitCompare = "SELSRAM_Compare_Bit_Str," & site & ","
            str_bitCompareAll = ""
            'str_bitDSSC = "SELSRAM_DSSC_Bit_Str," & site & "," & inst_info.str_Selsrm_DSSC_Bit(site)
            str_sram_vth = "SRAM_Vth(DCVS)," & site
            str_logic_voltage = "SelSram_voltage," & site
            
            For i = 0 To UBound(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).logic_Pin)
                powerDomain = UCase(Trim(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).logic_Pin(i)))
                
                '''//Check if powerDomain is "PRESERVED" or "RESERVED". C651 Toby requested this on 20200916.
                If powerDomain <> "PRESERVED" And powerDomain <> "RESERVED" Then
                    alphaPowerDomain = UCase(SelsramMapping(inst_info.idxBlock_Selsrm_singlePatt).alpha(i))
                    
                    str_bitCompare = str_bitCompare & bitArray_corePower(VddBinStr2Enum(powerDomain))
                    str_bitCompareAll = str_bitCompareAll & "," & alphaPowerDomain & "=" & bitArray_corePower(VddBinStr2Enum(powerDomain))
                    str_sram_vth = str_sram_vth & "," & alphaPowerDomain & "=" & CStr(Format(Floor(inst_info.sram_Vth(VddBinStr2Enum(powerDomain))) / 1000, "#0.000")) & "V"
                    str_logic_voltage = str_logic_voltage & "," & alphaPowerDomain & "=" & CStr(Format(Floor(voltage_SelsrmBitCalc(VddBinStr2Enum(powerDomain))) / 1000, "#0.000")) & "V"
                End If
            Next i
            'inst_info.str_Selsrm_DSSC_Info(site) = str_bitCompare & "(LSB->MSB)" & str_bitCompareAll & vbCrLf & str_bitDSSC & vbCrLf & str_sram_vth & vbCrLf & str_logic_voltage
            inst_info.str_Selsrm_DSSC_Info(site) = str_bitCompare & "(LSB->MSB)" & str_bitCompareAll & vbCrLf & str_sram_vth & vbCrLf & str_logic_voltage
        Next site
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Write_Selsrm_DSSC_Info"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Write_Selsrm_DSSC_Info"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200326: Created to decide if the powerbinning is for Harvest or not.
Public Function Decide_PowerBinning_Type()
    Dim idxSheet As Integer
    Dim cntSheetPwrbin As Integer
    Dim cntSheetPwrbinHarvest As Integer
    Dim strAry_sheetnames() As String
    Dim sheetnamePwrbin As String
On Error GoTo errHandler
    strAry_sheetnames = TheExec.Job.GetSheetNamesOfType(DMGR_SHEET_TYPE_USER)
    
    '''//Init variables
    cntSheetPwrbin = 0
    cntSheetPwrbinHarvest = 0
    sheetnamePwrbin = ""
    Flag_PowerBinningTable_Parsed = False
    
    For idxSheet = 0 To UBound(strAry_sheetnames)
        If LCase(strAry_sheetnames(idxSheet)) Like LCase("PwrBinning_V*") Then
            cntSheetPwrbinHarvest = cntSheetPwrbinHarvest + 1
            sheetnamePwrbin = strAry_sheetnames(idxSheet)
            
            If cntSheetPwrbinHarvest > 1 Then
                sheetnamePwrbin = ""
                TheExec.Datalog.WriteComment "PwrBinning_V* has the duplicate flow sequence sheet, Please check the workbook. Error!!!"
                TheExec.ErrorLogMessage "PwrBinning_V* has the duplicate flow sequence sheet, Please check the workbook. Error!!!"
            End If
        ElseIf LCase(strAry_sheetnames(idxSheet)) Like LCase("Pwrbin_Seq*") Then
            cntSheetPwrbin = cntSheetPwrbin + 1
            sheetnamePwrbin = strAry_sheetnames(idxSheet)
            
            If cntSheetPwrbinHarvest > 1 Then
                sheetnamePwrbin = ""
                TheExec.Datalog.WriteComment "Pwrbin_Seq* has the duplicate flow sequence sheet, Please check the workbook. Error!!!"
                TheExec.ErrorLogMessage "Pwrbin_Seq* has the duplicate flow sequence sheet, Please check the workbook. Error!!!"
            End If
        End If
    Next idxSheet
    
    If cntSheetPwrbinHarvest = 1 Then '''for powerBinning_Harvest
        Parsing_Power_Bin_Table_Harvest sheetnamePwrbin
    ElseIf cntSheetPwrbin = 1 Then
        Parsing_Power_Bin_Table sheetnamePwrbin
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Decide_PowerBinning_Type"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Decide_PowerBinning_Type"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210824: Modified to add the argument "PassBinNum As SiteLong" to the vbt function Calculate_Binning_CorePower_with_DynamicOffset.
'20210824: Modified to remove the redundant arguments "IndexLevelPerSite As SiteLong" and "voltage_forBinCut() As SiteDouble" from the vbt function Calculate_Binning_CorePower_with_DynamicOffset.
'20210824: Modified to move the vbt function Calculate_Binning_CorePower_with_DynamicOffset from calculate_payload_voltage_for_BV to GradeSearch_VT.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210125: Modified to move "Calculate_Extra_Voltage_for_PowerRail" from "Calculate_Binning_CorePower_with_DynamicOffset" to "calculate_payload_voltage_for_BV".
'20201215: Modified to reduce the redundant site-loop.
'20201210: Modified to remove "lvl_mv_dynamicOffset As SiteDouble" from the arguments of the vbt function Calculate_Binning_CorePower_with_DynamicOffset.
'20201204: Modified to add the argument "IndexLevelPerSite As SiteLong" for the vbt function Calculate_Binning_CorePower_with_DynamicOffset.
'20201116: Modified to use "testCondition = LCase(Trim(Get_BinCut_TestCondition()))".
'20201029: Modified to use "Public Type Instance_Info".
'20200809: Modified to truncate the EQN-based voltage of the binning CorePower.
'20200711: Modified to change the argumens types and store BinCut payload voltages into the siteDouble array.
'20200430: Modified to print the string about the incorrect voltage.
'20200429: Modified to use "Get_BinCut_TestCondition" to get BinCut testCondition.
'20200429: Modified to move this function from "LIB_VDD_BINNING" to "LIB_Vdd_Binning_customer".
'20200317: Modified for SearchByPmode.
'20200130: Modified to store binning CorePower voltage in the siteDouble array "BinCut_Payload_Voltage()".
'20200106: Modified to remove the ErrorLogMessage.
'20191227: Modified to check the string of testCondition in Non_Binning_Rail.
'20191224: Created to calculate binning corePower with dynamic offset.
'20190507: Modified to add "Cdec" to avoid double format accuracy issues.
'20190304: Modified to use Calculate_Extra_Voltage_for_PowerRail to calculate extra voltage for the binning power.
Public Function Calculate_Binning_CorePower_with_DynamicOffset(inst_info As Instance_Info, PassBinNum As SiteLong, voltage_binningPmode As SiteDouble)
    Dim site As Variant
    Dim dynamicOffset_BinCut As Double
    Dim bincutNum As New SiteLong
    Dim voltage_binningPmode_dynamicOffset As New SiteDouble
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Dynamic offset is not related to Efuse product voltage, so that it doesn't need take the least multiple of StepVoltage.
'''//==================================================================================================================================================================================//'''
    For Each site In TheExec.sites
        '''***********************************************************************************'''
        '''[Step1] Get Dynamic Offset of the binning CorePower.
        '''***********************************************************************************'''
        dynamicOffset_BinCut = BinCut(inst_info.p_mode, bincutNum(site)).DYNAMIC_OFFSET(inst_info.jobIdx, inst_info.offsetTestTypeIdx)
        
        '''***********************************************************************************'''
        '''[Step2] Calculate BinCut payload voltage with dynamicOffset.
        '''//Dynamic offset is not related to Efuse product voltage, so that it doesn't need take the least multiple of Step Size voltage.
        '''***********************************************************************************'''
        voltage_binningPmode_dynamicOffset = voltage_binningPmode + dynamicOffset_BinCut
                
        If dynamicOffset_BinCut <> 0 Then
            inst_info.str_dynamic_offset(site) = "Dynamic Offset," & site & "," & inst_info.powerDomain & ",Voltage Change From " & voltage_binningPmode & "mV To " & voltage_binningPmode_dynamicOffset & "mV (Offset=" & dynamicOffset_BinCut & "mV)"
        Else
            inst_info.str_dynamic_offset(site) = ""
        End If
              
        '''***********************************************************************************'''
        '''[Step3] Check if the binning CorePower voltage>0.
        '''***********************************************************************************'''
        If CDec(voltage_binningPmode_dynamicOffset) > 0 Then
            '''Do nothing
        Else
            voltage_binningPmode_dynamicOffset(site) = 0
            TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.powerDomain & ", BinCut payload voltage should be larger than 0. Error occurs in Calculate_Binning_CorePower_with_DynamicOffset while " & inst_info.inst_name & ". Error!!!"
            'TheExec.ErrorLogMessage "site:" & site & "," &  inst_info.powerDomain & ", BinCut payload voltage should be larger than 0. Error occurs in Calculate_Binning_CorePower_with_DynamicOffset while " & inst_info.inst_name & ". Error!!!"
        End If
    Next site
        
    '''//Store voltages of the binning CorePower into the siteDouble array "voltage_forBinCut".
    '''Due to the accuracy of DCVS is 1mv, it should truncate the EQN-based voltage of the binning CorePower.
    voltage_binningPmode = voltage_binningPmode_dynamicOffset.Truncate
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Calculate_Binning_CorePower_with_DynamicOffset"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210824: Modified to rename the vbt function calculate_payload_voltage_for_BV as get_passBin_from_Step.
'20210824: Modified to move the vbt function Non_Binning_Pwr_Setting_VT from calculate_payload_voltage_for_BV to GradeSearch_VT.
'20210824: Modified to move the vbt function Calculate_Binning_CorePower_with_DynamicOffset from calculate_payload_voltage_for_BV to GradeSearch_VT.
'20210729: Modified to remove the redundant argument "ids_current As SiteDouble" from the vbt functions Apply_testcondition_InFlowSheet and Calculate_LVCC_for_CorePower.
'20210728: Modified to move "Dim Sram_Vth(MaxBincutPowerdomainCount) As New SiteDouble" into "Public Type Instance_Info".
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210408: Modified to force step in Dynamic_IDS_Zone to step_interpolated if p_mode is interpolated = True And current PassBinNum = 1.
'20210407: Modified to revise the vbt code for the new Interpolation method proposed by C651 Toby.
'20210323: As per discussion with PCLINZG, modified to get passBinFromStep from step_Start in Dynamic_IDS_Zone for DevChar if p_mode is not tested.
'20210226: Modified to use step_Start and step_Stop to get startVoltage and StopVoltage.
'20210207: Modified to add the argument "idxPin_CorePower As Integer" to the vbt function Calculate_LVCC_for_CorePower and Calculate_HVCC_for_CorePower.
'20210201: Modified to remove the redundant vbt code of "set_BinCut_voltage_for_DevChar".
'20210126: Modified to revise the vbt code for DevChar.
'20210125: Modified to remove "voltage_Pmode_EQNbased As SiteDouble" from the arguments of the vbt function "calculate_payload_voltage_for_BV".
'20210125: Modified to move "Calculate_Extra_Voltage_for_PowerRail" from "Calculate_Binning_CorePower_with_DynamicOffset" to "calculate_payload_voltage_for_BV".
'20201210: Modified to remove "lvl_mv_dynamicOffset As SiteDouble" from the arguments of the vbt functions calculate_payload_voltage_for_BV and Calculate_Binning_CorePower_with_DynamicOffset.
'20201210: Modified to rename the vbt function "calculate_payload_voltage_for_binning_CorePower" as "calculate_payload_voltage_for_BV".
'20201210: Modified to add the vbt functions "Get_PassBinNum_by_Step" and "Non_Binning_Pwr_Setting_VT" into the vbt function calculate_payload_voltage_for_binning_CorePower.
'20201210: Modified to use the arguments "inst_info As Instance_Info" for Get_Binning_CorePower_PayloadVoltage_by_Step.
'20201207: Created to calculate payload with Dynamic Offset for the binning CorePower.
Public Function get_passBin_from_Step(inst_info As Instance_Info, passBinFromStep As SiteLong)
    Dim site As Variant
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Get passBin by step in Dynamic IDS Zone of the binning performance mode.
'''//==================================================================================================================================================================================//'''
    If inst_info.is_DevChar_Running = True Then '''for DevChar.
        For Each site In TheExec.sites
            If VBIN_RESULT(inst_info.p_mode).tested = True Then '''//Check if p_mode is tested.
                passBinFromStep = VBIN_RESULT(inst_info.p_mode).passBinCut
            Else
                '''20210323: As per discussion with PCLINZG, modified to get passBinFromStep from step_Start in Dynamic_IDS_Zone for DevChar if p_mode is not tested.
                '''//step_control.step_Start(site) is updated in the vbt function "find_start_voltage" according to Dynamic_IDS_Zone and CurrentPassBinCutNum from previous test instances.
                passBinFromStep = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Start(site))
            End If
        Next site
    Else '''If inst_info.is_DevChar_Running = False
        If inst_info.is_BinSearch = True Then '''BinCut search.
            '=====================================================================================
            '[Step1] Get EQN-based BinCut payload voltage of the binning corePower by step in Dynamic IDS Zone.
            '=====================================================================================
            For Each site In TheExec.sites
                '''//If p_mode is interpolated, it needs to check step in Dynamic_IDS_Zone.
                '''***************************************************************************************************************************************************************************************'''
                '''//Only Bin1 DUT uses Interpolation, and Bin1 has the Interpolation factor.
                '''//If p_mode is interpolated:
                '''1. current step in Dynamic_IDS_Zone is lower than DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Interpolated_Start, it can overwrite current step (inst_info.IndexLevelPerSite) to DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Interpolated_Start for Interpolation.
                '''2. current step in Dynamic_IDS_Zone is greater than DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Interpolated_Start, it can overwrite current step (inst_info.IndexLevelPerSite) to DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_mapping(passbinBinCut(inst_info.IndexLevelPerSite),EQN1) for Interpolation.
                '''20201113: C651 Toby requested the new method to check BinCut voltages montonicity, so that we need to keep all available steps in Dynamic_IDS_Zone.
                '''20210408: Modified to force step in Dynamic_IDS_Zone to step_interpolated if p_mode is interpolated = True And current PassBinNum = 1.
                '''***************************************************************************************************************************************************************************************'''
                If DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).interpolated = True And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current) = 1 Then
                    If inst_info.step_Current(site) <= DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Interpolated_Start Then
                        inst_info.step_Current = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Interpolated_Start
                    ElseIf inst_info.step_Current(site) <= DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current), 1) Then
                        inst_info.step_Current(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current), 1)
                    End If
                End If
                
                '''//Get PassBinNum by Step in Dynamic IDS Zone of the binning p_mode.
                passBinFromStep(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current)
            Next site
        Else '''BinCut check.
            '''//Get PassBinNum from CurrentPassBinCutNum(the fused PassBin).
            passBinFromStep = CurrentPassBinCutNum
        End If
    End If '''If inst_info.is_DevChar_Running = True
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of get_passBin_from_Step"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210129: Modified to revise the vbt code for DevChar.
'20210126: Modified to revise the vbt code for DevChar.
Public Function set_BinCut_voltage_for_DevChar(ByRef voltage_forBinCut() As SiteDouble)
    Dim site As Variant
    '''for DevChar
    Dim ShmooPinsIist As String
    Dim ShmooPinsArr() As String
    Dim ShmooPinsCount As Integer
    Dim ShmooPin As String
On Error GoTo errHandler
    '''//Set g_Vbump_function = True to get Shmoo voltage settings (Shmoo conditions) by "Get_Current_Apply_Pin".
    g_Vbump_function = True
    
    '''//Catch shmoo voltage and override to Bincut_Payload_Voltage.
    Call Get_Current_Apply_Pin(ShmooPinsIist)
    
    '''//After getting Shmoo voltage settings (Shmoo conditions), remember to set g_Vbump_function = False.
    g_Vbump_function = False
    
    ShmooPinsArr = Split(ShmooPinsIist, ",")
    
    For ShmooPinsCount = 0 To UBound(ShmooPinsArr)
        ShmooPin = ShmooPinsArr(ShmooPinsCount)
        For Each site In TheExec.sites
            voltage_forBinCut(VddBinStr2Enum(ShmooPin))(site) = g_Globalpointval.Pins(ShmooPin).Value * 1000
        Next site
    Next
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of calculate_payload_voltage_for_binning_CorePower"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201029: Modified to replace the argument "pmode_of_powerDomain as Integer" with "idxPin_PowerDomain As Integer".
'20201029: Modified to use "Public Type Instance_Info".
'20200702: Modified to check the instance keyword "outsidebincut".
'20200429: Created to get testCondition by the keyword "_BV" or "_BV" of the instance and PassBinNum.
'20191227: Modified to check the string of Non_Binning_Rail.
'20190312: Modified to check if the instance is bv or hbv.
Public Function Get_BinCut_TestCondition(inst_info As Instance_Info, idxPin_PowerDomain As Integer, PassBinNum As Long) As String
    Dim strTemp As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''If the instance name contains the keyword "*_outsidebincut_*", get testCondition from OutsideBinCut table.
'''//==================================================================================================================================================================================//'''
    '''//Get the instance name.
    strTemp = LCase(inst_info.inst_name)
    
    '''//Check if the instance is bv or hbv, then get testCondition from Non_Binning_Rail for CorePower and OtherRail.
    '''Check if the instance keyword "outsidebincut" exists for Outside BinCut test instances.
    If strTemp Like "*_hbv" Then 'hbv test instance
        If inst_info.special_voltage_setup = True Then
            If strTemp Like "*_outsidebincut_*" Then
                Get_BinCut_TestCondition = LCase(Trim(BinCut(inst_info.p_mode, PassBinNum).OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE(idxPin_PowerDomain, inst_info.addi_mode)))
            Else
                Get_BinCut_TestCondition = LCase(Trim(BinCut(inst_info.p_mode, PassBinNum).HVCC_Addtional_OTHER_VOLTAGE(idxPin_PowerDomain, inst_info.addi_mode)))
            End If
        Else
            If strTemp Like "*_outsidebincut_*" Then
                Get_BinCut_TestCondition = LCase(Trim(BinCut(inst_info.p_mode, PassBinNum).OutsideBinCut_HVCC_OTHER_VOLTAGE(idxPin_PowerDomain)))
            Else
                Get_BinCut_TestCondition = LCase(Trim(BinCut(inst_info.p_mode, PassBinNum).HVCC_OTHER_VOLTAGE(idxPin_PowerDomain)))
            End If
        End If
    ElseIf strTemp Like "*_bv" Then 'bv test instance
        If inst_info.special_voltage_setup = True Then
            If strTemp Like "*_outsidebincut_*" Then
                Get_BinCut_TestCondition = LCase(Trim(BinCut(inst_info.p_mode, PassBinNum).OutsideBinCut_Addtional_OTHER_VOLTAGE(idxPin_PowerDomain, inst_info.addi_mode)))
            Else
                Get_BinCut_TestCondition = LCase(Trim(BinCut(inst_info.p_mode, PassBinNum).Addtional_OTHER_VOLTAGE(idxPin_PowerDomain, inst_info.addi_mode)))
            End If
        Else
            If strTemp Like "*_outsidebincut_*" Then
                Get_BinCut_TestCondition = LCase(Trim(BinCut(inst_info.p_mode, PassBinNum).OutsideBinCut_OTHER_VOLTAGE(idxPin_PowerDomain)))
            Else
                Get_BinCut_TestCondition = LCase(Trim(BinCut(inst_info.p_mode, PassBinNum).OTHER_VOLTAGE(idxPin_PowerDomain)))
            End If
        End If
    Else
        TheExec.Datalog.WriteComment "Instance: " & inst_info.inst_name & " doesn't end with any _bv or _hbv keyword. Error!!!"
        'TheExec.ErrorLogMessage "instance: " & inst_info.inst_name & " doesn't end with any _bv or _hbv keyword. Error!!!"
    End If
    
    '''//Check if testCondition from Non_Binning_Rail is empty.
    If Get_BinCut_TestCondition = "" Or Get_BinCut_TestCondition = "0" Then
        TheExec.Datalog.WriteComment "PassBinNum:" & PassBinNum & ", " & VddBinName(inst_info.p_mode) & ", selected powerDomain: " & VddBinName(idxPin_PowerDomain) & " has incorrect content in the Non_Binning_Rail. Error!!!"
        'TheExec.ErrorLogMessage "PassBinNum:" & PassBinNum & ", " & VddBinName(inst_info.p_mode) & ", selected powerDomain: " & VddBinName(idxPin_PowerDomain) & " has incorrect content in the Non_Binning_Rail. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Get_BinCut_TestCondition"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Get_BinCut_TestCondition"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210906: Modified to check if the current testJob is for BinCut search.
'20210325: Modified to use Flag_Vddbin_DoAll_DebugCollection for TheExec.EnableWord("Vddbin_DoAll_DebugCollection").
'20210322: Modified to decide Flag_Enable_CMEM_Collection by checking TheExec.Flow.EnableWord("Vddbin_CMEM_Collection").
'20210322: Modified to decide Flag_Vddbin_COF_Instance and Vddbin_COF_Instance_with_PerEqnLog by checking TheExec.Flow.EnableWord("Vddbin_COF_Instance") and TheExec.Flow.EnableWord("Vddbin_COF_Instance_with_PerEqnLog").
'20210322: Modified to decide Flag_Vddbin_COF_StepInheritance by checking TheExec.Flow.EnableWord("Vddbin_COF_StepInheritance").
'20210113: Modified to mask the vbt code of checking the flags "VDDBinning_Offline_AllPattPass" and "Golden_Default", requested by Leon Weng.
'20201126: Modified to check if Flag_Enable_CMEM_Collection = True while Opensocket or Offline.
'20201020: Modified to check "COFInstance" and "PerEqnLog".
'20200903: Modified to check if "Vddbin_DoAll_DebugCollection" is enabled for non-CP1.
'20200807: Modified to set "Flag_Enable_CMEM_Collection=False" if the tester is opensocket or IGXL is offline.
'20200730: Modified to check "Golden_Default" and "VDDBinning_Offline_AllPattPass".
'20200527: Created to check BinCut flags and globalVariables.
'20190905: Modified the branch that FFC CMEM is only for online test.
Public Function Check_BinCut_flag_globalVariable()
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Check if any conflict of the flag settings with Opensocket or Offline.
'''//==================================================================================================================================================================================//'''
    '''//Check if Vbt settings and the testMode are offline or not.
    If TheExec.TesterMode = testModeOffline Then '''The tester is offline.
        Flag_VDD_Binning_Offline = True
    Else '''The tester is online.
        If Flag_Tester_Offline = True Then '''The tester is online, but IGXL is forced to offline.
            Flag_VDD_Binning_Offline = True
        Else '''IGXL is online.
            Flag_VDD_Binning_Offline = False
        End If
    End If

    '''//Check if Flag_Enable_CMEM_Collection = True while Opensocket or Offline.
    '''If the tester is opensocket or offline, skip using CMEM.
    If Flag_VDD_Binning_Offline = True Or EnableWord_Vddbinning_OpenSocket = True Then
        If TheExec.Flow.EnableWord("Vddbin_CMEM_Collection") = True Then
            '''//For offline or OpenSocket, CMEM_Collection can't be activated.
            TheExec.Flow.EnableWord("Vddbin_CMEM_Collection") = False
            '''//Show the error message.
            TheExec.Datalog.WriteComment "Flag_Enable_CMEM_Collection is enabled, but it doesn't support CMEM while Opensocket or Offline. Please check the flag setting. Error!!!"
            'TheExec.ErrorLogMessage "Flag_Enable_CMEM_Collection is enabled, but it doesn't support CMEM while Opensocket or Offline. Please check the flag setting. Error!!!"
        End If
    Else
        '''//Set default value for the flag TheExec.Flow.EnableWord("Vddbin_CMEM_Collection").
        TheExec.Flow.EnableWord("Vddbin_CMEM_Collection") = True
    End If
    
    '''//Check TheExec.Flow.EnableWord("Vddbin_CMEM_Collection").
    If TheExec.Flow.EnableWord("Vddbin_CMEM_Collection") = True Then
        Flag_Enable_CMEM_Collection = True
    Else
        Flag_Enable_CMEM_Collection = False
    End If
    
    '''//Check "COFInstance" and "PerEqnLog" for COF_Instance.
    If TheExec.Flow.EnableWord("Vddbin_COF_Instance") = True Or TheExec.Flow.EnableWord("Vddbin_COF_Instance_with_PerEqnLog") = True Then
        If Flag_IDS_Distribution_enable = True Then
            Flag_Vddbin_COF_Instance = False
            Flag_Vddbin_COF_Instance_with_PerEqnLog = False
            TheExec.Datalog.WriteComment "Flag_IDS_Distribution_enable is True, and Flag_Vddbin_COF_Instance is True. Error!!!"
            TheExec.ErrorLogMessage "Flag_IDS_Distribution_enable is True, and Flag_Vddbin_COF_Instance is True. Error!!!"
        Else
            Flag_Vddbin_COF_Instance = True
            If TheExec.Flow.EnableWord("Vddbin_COF_Instance_with_PerEqnLog") = True Then
                Flag_Vddbin_COF_Instance_with_PerEqnLog = True
            Else
                Flag_Vddbin_COF_Instance_with_PerEqnLog = False
            End If
        End If
    Else
        Flag_Vddbin_COF_Instance = False
        Flag_Vddbin_COF_Instance_with_PerEqnLog = False
    End If
    
    '''//Check "COF_StepInheritance".
    If TheExec.Flow.EnableWord("Vddbin_COF_StepInheritance") = True Then
        Flag_Vddbin_COF_StepInheritance = True
    Else
        Flag_Vddbin_COF_StepInheritance = False
    End If

    '''//Check "Vddbin_DoAll_DebugCollection" with "Vddbin_COF_StepInheritance".
    '''20210325: Modified to use Flag_Vddbin_DoAll_DebugCollection for TheExec.EnableWord("Vddbin_DoAll_DebugCollection").
    If TheExec.Flow.EnableWord("Vddbin_COF_StepInheritance") = True And TheExec.EnableWord("Vddbin_DoAll_DebugCollection") = True Then
        Flag_Vddbin_DoAll_DebugCollection = False
        TheExec.Datalog.WriteComment "EnableWords Vddbin_COF_StepInheritance and Vddbin_DoAll_DebugCollection are enabled, but it makes BinCut GradeSearch rules with conflicts. Error!!!"
        TheExec.ErrorLogMessage "EnableWords Vddbin_COF_StepInheritance and Vddbin_DoAll_DebugCollection are enabled, but it makes BinCut GradeSearch rules with conflicts. Error!!!"
    Else
        If TheExec.EnableWord("Vddbin_DoAll_DebugCollection") = True Then
            Flag_Vddbin_DoAll_DebugCollection = True
        Else
            Flag_Vddbin_DoAll_DebugCollection = False
        End If
    End If
    
    '''//Check if "Vddbin_DoAll_DebugCollection" is enabled for BinCut check.
    '''20210906: Modified to check if the current testJob is for BinCut search.
    If is_BinCutJob_for_StepSearch = False And TheExec.EnableWord("Vddbin_DoAll_DebugCollection") = True Then
        TheExec.Datalog.WriteComment "Vddbin_DoAll_DebugCollection is only for testJob with BinCut search. Please check EnableWord settings. Error!!!"
        TheExec.ErrorLogMessage "Vddbin_DoAll_DebugCollection is only for testJob with BinCut search. Please check EnableWord settings. Error!!!"
        TheExec.EnableWord("Vddbin_DoAll_DebugCollection") = False
    End If
    
'    '''//**********Preserve this area for checking the rules of flags in "LIB_Vdd_Binning_GlobalVariable"**********//'''
'    If Flag_SyncUp_DCVS_Output_enable Then
'        If Flag_Enable_Rail_Switch Then
'            '''Do nothing
'        Else
'            TheExec.Datalog.WriteComment "Flag_SyncUp_DCVS_Output_enable is enabled, but Flag_Enable_Rail_Switch is disabled. Please check the flags setting. Error!!!"
'            TheExec.ErrorLogMessage "Flag_SyncUp_DCVS_Output_enable is enabled, but Flag_Enable_Rail_Switch is disabled. Please check the flags setting. Error!!!"
'        End If
'    End If
'
'    If Flag_noRestoreVoltageForPrepatt = True Then
'        If Flag_Enable_Rail_Switch Then
'            '''Do nothing
'        Else
'            TheExec.Datalog.WriteComment "Flag_noRestoreVoltageForPrepatt is enabled, but Flag_Enable_Rail_Switch is disabled. Please check the flags setting. Error!!!"
'            TheExec.ErrorLogMessage "Flag_noRestoreVoltageForPrepatt is enabled, but Flag_Enable_Rail_Switch is disabled. Please check the flags setting. Error!!!"
'        End If
'    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Check_BinCut_flag_globalVariable"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Check_BinCut_flag_globalVariable"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200609: Created to check alarmFail before BinCut Initial.
'20190524: Modified to monitor alarmFail.
Public Function Check_alarmFail_before_BinCut_Initial(Optional inst_name As String = "")
    Dim site As Variant
On Error GoTo errHandler
    For Each site In TheExec.sites
        If alarmFail(site) = True Then
            If inst_name <> "" Then
                TheExec.Datalog.WriteComment "site:" & site & ", alarmFail is triggered before the instance: " & inst_name & ". alarmFail!!!"
            Else
                TheExec.Datalog.WriteComment "site:" & site & ", alarmFail!!!"
            End If
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Check_alarmFail_before_BinCut_Initial"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Check_alarmFail_before_BinCut_Initial"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200917: Modified to merge the branches of bv and hbv.
'20200622: Modified for the flag naming rule for the failFlag of Call Instance.
'20200615: Created to get failflagName by the instance name.
Public Function Get_flagName_from_instanceName(inst_name As String, p_mode As Integer, flagName As String)
    Dim strTemp As String
On Error GoTo errHandler
    strTemp = LCase(inst_name)
    
    If strTemp Like "*_bv" Or strTemp Like "*_hbv" Then '''BinCut BV or HBV instance.
        flagName = "F_BV_CALLINST"
    Else
        flagName = "F_NoCorrectInstance"
        TheExec.Datalog.WriteComment "The instance name " & inst_name & " is incorrect for the function Get_flagName_from_instanceName. Error!!!"
        TheExec.ErrorLogMessage "The instance name " & inst_name & " is incorrect for the function Get_flagName_from_instanceName. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Get_flagName_from_instanceName"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Get_flagName_from_instanceName"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201102: Modified to add the argument "enable_DecomposePatt as boolean"
'20201014: Modified to revise the vbt code of checking payload patterns, requested by Leon Weng.
'20201007: Modified to check strAry_PatNameSplit(3) and strAry_PatNameSplit(4). requested by Leon Weng.
'20200809: Created to check DCVS output and Payload pattern.
'20200802: Modified to check patType init or payload, revised by Leon.
'20200527: Modified to check Vbump for project with rail-switch.
'20200217: Modified to check if no vbump before running the payload pattern.
Public Function Check_PayloadPattern_with_DCVS(inst_name As String, flagEnableRailSwitch As Boolean, PatName As String, currentDcvsOutput As Integer, enable_DecomposePatt As Boolean)
    Dim strAry_PatNameSplit() As String
    Dim strAry_PathSplit() As String
    Dim hasKeywordPayload As Boolean
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''DCVS should be switched to Valt by Pattern with vbump prior to Payload pattern for TD test instances.
'''//==================================================================================================================================================================================//'''
    If flagEnableRailSwitch Then
        '''init
        strAry_PathSplit = Split(LCase(PatName), "\")
        strAry_PatNameSplit = Split(LCase(strAry_PathSplit(UBound(strAry_PathSplit))), "_")
        hasKeywordPayload = False
        
        '''ToDo: Maybe we can check patterns with keywords "*_pllp*", "*_fulp*", and "*_pl*"...
        If strAry_PatNameSplit(3) Like "*pl*" Then
            hasKeywordPayload = True
        End If
        
        If enable_DecomposePatt = True And (LCase(inst_name) Like "*td*") And currentDcvsOutput = tlDCVSVoltageMain And hasKeywordPayload = True Then
            TheExec.Datalog.WriteComment "Warning!!! No vbump to select DCVS Valt before the payload pattern. Please check the pattern set. Warning!!!"
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Check_PayloadPattern_with_DCVS"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Check_PayloadPattern_with_DCVS"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210611: Modified to update Flag_Harvest_Core_DSSC_Ready = True if Harv_Pmode_Table and HarvMappingTable were parsed.
'20210518: Modified to check TheExec.CurrentJob for worksheet.
'20210511: Modified to all related dictionaries for Harvest Core DSSC.
'20210510: Modified to parse HARV_Pmode_Table and HARVMappingTable for Harvest Core DSSC.
Public Function check_Harvest_Core_All_Table()
    Dim strAry_sheetnames() As String
    Dim idxSheet As Integer
    Dim cntSheet_Harvest_Pmode As Integer
    Dim cntSheet_Harvest_Mapping As Integer
    Dim sheetname_Harvest_Pmode As String
    Dim sheetname_Harvest_Mapping As String
On Error GoTo errHandler
    '''//"Harv_Pmode_Table" and "HARVMappingTable" are user-defined sheets of IGXL, so that sheet type should be DMGR_SHEET_TYPE_USER.
    strAry_sheetnames = TheExec.Job.GetSheetNamesOfType(DMGR_SHEET_TYPE_USER)
    
    '''//Init variables
    Flag_Harvest_Pmode_Table_Parsed = False
    Flag_Harvest_Mapping_Table_Parsed = False
    Flag_Harvest_Core_DSSC_Ready = False
    cntSheet_Harvest_Pmode = 0
    cntSheet_Harvest_Mapping = 0
    sheetname_Harvest_Pmode = ""
    sheetname_Harvest_Mapping = ""
    
    '''//Reset all Harve Core DSSC related dictionaries before parsing Harvest Core DSSC related tables.
    '''20210511: Modified to all related dictionaries for Harvest Core DSSC.
    dict_Pmode2ByMode.RemoveAll
    dict_ByMode2Index.RemoveAll
    dict_DisableCore2FailFlag.RemoveAll
    dict_FailFlag2DisableCore.RemoveAll
    dict_FailFlagOfDisableCore2DevCondition.RemoveAll
    dict_HarvestCoreGroup2Index.RemoveAll
    dict_EnableCore2Fstp.RemoveAll
    dict_Fstp2EnableCore.RemoveAll
    
    For idxSheet = 0 To UBound(strAry_sheetnames)
        '''//SheetName should contain the keyword about testJob.
        '''20210518: Modified to check TheExec.CurrentJob for worksheet.
        If LCase(strAry_sheetnames(idxSheet)) Like LCase("HARV_Pmode_Table_" & TheExec.CurrentJob) Then '''ex: "HARV_Pmode_Table_CP1"
            cntSheet_Harvest_Pmode = cntSheet_Harvest_Pmode + 1
            sheetname_Harvest_Pmode = strAry_sheetnames(idxSheet)
            
            If cntSheet_Harvest_Pmode > 1 Then
                sheetname_Harvest_Pmode = ""
                TheExec.Datalog.WriteComment "sheet HARV_Pmode_Table has the duplicate flow sequence sheet, Please check the workbook for check_Harvest_Core_All_Table. Error!!!"
                TheExec.ErrorLogMessage "sheet HARV_Pmode_Table has the duplicate flow sequence sheet, Please check the workbook for check_Harvest_Core_All_Table. Error!!!"
            End If
        ElseIf LCase(strAry_sheetnames(idxSheet)) Like LCase("HARVMappingTable_" & TheExec.CurrentJob) Then '''ex: "HARVMappingTable_CP1".
            cntSheet_Harvest_Mapping = cntSheet_Harvest_Mapping + 1
            sheetname_Harvest_Mapping = strAry_sheetnames(idxSheet)
            
            If cntSheet_Harvest_Mapping > 1 Then
                sheetname_Harvest_Mapping = ""
                TheExec.Datalog.WriteComment "sheet HARVMappingTable has the duplicate flow sequence sheet, Please check the workbook for check_Harvest_Core_All_Table. Error!!!"
                TheExec.ErrorLogMessage "sheet HARVMappingTable has the duplicate flow sequence sheet, Please check the workbook for check_Harvest_Core_All_Table. Error!!!"
            End If
        End If
    Next idxSheet
    
    '''//If parsing Harvest Core DSSC related tables are available in the workbook, start to parse these tables.
    If cntSheet_Harvest_Pmode > 0 And cntSheet_Harvest_Mapping > 0 Then '''for powerBinning_Harvest
        '''//Parse HARV_Pmode_Table.
        Parsing_Harvest_Pmode_Table sheetname_Harvest_Pmode
        
        '''//Parse HARVMappingTable.
        If Flag_Harvest_Pmode_Table_Parsed = True Then
            Parsing_Harvest_Mapping_Table sheetname_Harvest_Mapping
            
            '''//Check results.
            If Flag_Harvest_Mapping_Table_Parsed = True Then
                '''20210611: Modified to update Flag_Harvest_Core_DSSC_Ready = True if Harv_Pmode_Table and HarvMappingTable were parsed.
                Flag_Harvest_Core_DSSC_Ready = True
            End If
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of check_Harvest_Core_All_Table"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of check_Harvest_Core_All_Table"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210517: Modified to check instance keywords, ex: "MGX001".
'20210510: Modified to parse HARV_Pmode_Table and HARVMappingTable for Harvest Core DSSC.
Public Function Parsing_Harvest_Pmode_Table(sheetName As String)
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim isSheetFound As Boolean
    Dim row As Long, col As Long
    '''flags
    Dim enable_Parsing_Pmode2ByMode As Boolean
    Dim enable_Parsing_FailFlag2DiableCore As Boolean
    Dim enable_Parsing_ByMode2CoreGroup As Boolean
    Dim enable_Parsing_EnableCore2Fstp As Boolean
    '''pmode2ByMode
    Dim performance_mode As String
    Dim p_mode As Integer
    Dim str_ByMode As String
    Dim cnt_ByMode As Integer
    '''failflag2DisableCore
    Dim str_FailFlag As String
    Dim str_DiableCore As String
    '''Groups
    Dim cnt_HarvestCoreGroup As Integer
    Dim dict_columnOfTable As New Dictionary
    Dim bit_overWriteSeq_start As Long
    Dim bit_overWriteSeq_stop As Long
    Dim split_content() As String
    Dim str_Temp As String
    Dim str_prv_ByMode As String
    Dim cnt_Condition_of_ByMode As Long
    Dim idx_ByMode As Long
    Dim idx_GroupName As Long
    Dim str_GroupName As String
    Dim str_overWriteSeq As String
    Dim cnt_overWriteSeq As Long
    Dim lngAry_overWriteSeq() As Long
    Dim lngAry_coreGroup() As Long
    Dim i As Long
    '''EnableCore2Fstp
    Dim str_EnableCore As String
    Dim str_Fstp As String
    Dim cnt_Fstp As Long
On Error GoTo errHandler
    '''*****************************************************************'''
    '''//Check if the sheet exists
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        enable_Parsing_Pmode2ByMode = False
        enable_Parsing_FailFlag2DiableCore = False
        enable_Parsing_ByMode2CoreGroup = False
        enable_Parsing_EnableCore2Fstp = False
        bit_overWriteSeq_start = -1
        bit_overWriteSeq_stop = -1
        cnt_Condition_of_ByMode = 0
        
        For row = 1 To MaxRow
            '''init
            str_ByMode = ""
            str_FailFlag = ""
            str_DiableCore = ""
            str_EnableCore = ""
            str_Fstp = ""
            str_overWriteSeq = ""
            cnt_overWriteSeq = 0
            
            '''*****************************************************************'''
            '''//Find headers of each block.
            '''*****************************************************************'''
            If LCase(ws_def.Cells(row, 1).Value) Like LCase("Pmode") And LCase(ws_def.Cells(row, 2).Value) Like LCase("By_Mode") Then
                cnt_ByMode = 0
                row = row + 1
                enable_Parsing_Pmode2ByMode = True
                enable_Parsing_FailFlag2DiableCore = False
                enable_Parsing_ByMode2CoreGroup = False
                enable_Parsing_EnableCore2Fstp = False
                Flag_Harvest_Pmode_Table_Parsed = True
                
            ElseIf LCase(ws_def.Cells(row, 1).Value) Like LCase("HARV_Check") And LCase(ws_def.Cells(row, 3).Value) Like LCase("Disable_Core") Then
                row = row + 1
                enable_Parsing_Pmode2ByMode = False
                enable_Parsing_FailFlag2DiableCore = True
                enable_Parsing_ByMode2CoreGroup = False
                enable_Parsing_EnableCore2Fstp = False
                Flag_Harvest_Pmode_Table_Parsed = True
                
            ElseIf LCase(ws_def.Cells(row, 1).Value) Like LCase("By_Mode") And LCase(ws_def.Cells(row, 2).Value) Like LCase("Main_Core") Then
                cnt_HarvestCoreGroup = 0
                dict_columnOfTable.RemoveAll
                
                '''//Check if Group1 / Group2 / Group3 exist in the header.
                For col = 1 To maxcol
                    If ws_def.Cells(row, col).Value <> "" Then
                        '''//If core_group exists in dict_columnOfTable, it means that duplicate groups in the header.
                        If dict_columnOfTable.Exists(UCase(Trim(ws_def.Cells(row, col).Value))) = True Then
                            enable_Parsing_Pmode2ByMode = False
                            enable_Parsing_FailFlag2DiableCore = False
                            enable_Parsing_ByMode2CoreGroup = False
                            enable_Parsing_EnableCore2Fstp = False
                            TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & "," & col & "), it has the duplicate column for Parsing_Harvest_Pmode_Table. Error!!!"
                            TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & "," & col & "), it has the duplicate column for Parsing_Harvest_Pmode_Table. Error!!!"
                        Else
                            '''//Add core_group to dict_columnOfTable.
                            dict_columnOfTable.Add UCase(Trim(ws_def.Cells(row, col).Value)), col
                            
                            If UCase(Trim(ws_def.Cells(row, col).Value)) Like UCase("Group*") Then
                                If dict_HarvestCoreGroup2Index.Exists(UCase(Trim(ws_def.Cells(row, col).Value))) = True Then
                                    enable_Parsing_Pmode2ByMode = False
                                    enable_Parsing_FailFlag2DiableCore = False
                                    enable_Parsing_ByMode2CoreGroup = False
                                    enable_Parsing_EnableCore2Fstp = False
                                    TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & "," & col & "), it has the duplicate Group for Parsing_Harvest_Pmode_Table. Error!!!"
                                    TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & "," & col & "), it has the duplicate Group for Parsing_Harvest_Pmode_Table. Error!!!"
                                    Exit For
                                Else
                                    dict_HarvestCoreGroup2Index.Add UCase(Trim(ws_def.Cells(row, col).Value)), cnt_HarvestCoreGroup
                                    ReDim Preserve strAry_HarvestCoreGroupName(cnt_HarvestCoreGroup)
                                    strAry_HarvestCoreGroupName(cnt_HarvestCoreGroup) = UCase(Trim(ws_def.Cells(row, col).Value))
                                    cnt_HarvestCoreGroup = cnt_HarvestCoreGroup + 1
                                End If
                            '''//Get bit position of the replacement bits in Harvest Core DSSC bit sequence.
                            '''For example, Sequence[10:13] will be used to replace bit10~13 in Harvest Core DSSC bit sequence.
                            ElseIf UCase(Trim(ws_def.Cells(row, col).Value)) Like UCase("Sequence*#*#*") Then '''ex:Sequence[10:13]
                                '''//Get bit_overWriteSeq_start and bit_overWriteSeq_stop from Sequence[10:13].
                                split_content = Split(UCase(Trim(ws_def.Cells(row, col).Value)), "[")
                                '''//Sequence
                                dict_columnOfTable.Add UCase(Trim(split_content(0))), col
                                '''//Remove "]"
                                str_Temp = split_content(1)
                                split_content = Split(str_Temp, "]")
                                str_Temp = split_content(0)
                                split_content = Split(str_Temp, ":")
                                bit_overWriteSeq_start = CLng(Trim(split_content(0)))
                                bit_overWriteSeq_stop = CLng(Trim(split_content(1)))
                            End If
                        End If
                    Else
                        Exit For
                    End If
                Next col
                
                '''//If that, enable the flag to parse the table.
                row = row + 1
                enable_Parsing_Pmode2ByMode = False
                enable_Parsing_FailFlag2DiableCore = False
                enable_Parsing_ByMode2CoreGroup = True
                enable_Parsing_EnableCore2Fstp = False
                Flag_Harvest_Pmode_Table_Parsed = True
                
                '''
                ReDim Preserve HarvCoreByMode(dict_ByMode2Index.Count - 1)
                str_ByMode = ""
                str_prv_ByMode = ""
                cnt_Condition_of_ByMode = 0
                
            ElseIf LCase(ws_def.Cells(row, 1).Value) Like LCase("Enable_Core") And LCase(ws_def.Cells(row, 2).Value) Like LCase("FSTP") Then
                cnt_Fstp = 0
                row = row + 1
                enable_Parsing_Pmode2ByMode = False
                enable_Parsing_FailFlag2DiableCore = False
                enable_Parsing_ByMode2CoreGroup = False
                enable_Parsing_EnableCore2Fstp = True
                Flag_Harvest_Pmode_Table_Parsed = True
                
            Else
                If LCase(ws_def.Cells(row, 1).Value) <> "end" Then
                    If Trim(ws_def.Cells(row, 1).Value) = "" Then
                        enable_Parsing_Pmode2ByMode = False
                        enable_Parsing_FailFlag2DiableCore = False
                        enable_Parsing_ByMode2CoreGroup = False
                        enable_Parsing_EnableCore2Fstp = False
                    End If
                Else
                    Exit For
                End If
            End If
            
            '''*****************************************************************'''
            '''//Parse each block into dictionaries.
            '''*****************************************************************'''
            If enable_Parsing_Pmode2ByMode = True Then
                '''//Store the pattern_pmode related to keyword in the instance name or the pattern.
                '''For example, "MGX001" is the pattern_pmode used to check the instance name "MG001_GpuTd_MGX001_X4_Group1_PP_STNA0_L_PL00_PL01_PL02_PL03_PL04_SC_CFXX_TDF_COM_AUT_MGXXXX_DM_MLTFSTP_XOR_BV".
                '''20210517: Modified to check instance keywords, ex: "MGX001".
                performance_mode = UCase(Trim(ws_def.Cells(row, 1).Value))
                
                '''//Check if pmode exists in dict_Pmode2ByMode.
                If enable_Parsing_Pmode2ByMode = True Then
                    If dict_Pmode2ByMode.Exists(performance_mode) = True Then
                        enable_Parsing_Pmode2ByMode = False
                        Flag_Harvest_Pmode_Table_Parsed = False
                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & ",1), Pmode:" & ws_def.Cells(row, 1).Value & " has the duplicate definition in dict_Pmode2ByMode for Parsing_Harvest_Pmode_Table. Error!!!"
                        TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & ",1), Pmode:" & ws_def.Cells(row, 1).Value & " has the duplicate definition in dict_Pmode2ByModefor Parsing_Harvest_Pmode_Table. Error!!!"
                    Else
                        '''//Check if ByMode is correct.
                        '''For example, "X4" is the By_Mode used to check the instance name "MG001_GpuTd_MGX001_X4_Group1_PP_STNA0_L_PL00_PL01_PL02_PL03_PL04_SC_CFXX_TDF_COM_AUT_MGXXXX_DM_MLTFSTP_XOR_BV".
                        If LCase(ws_def.Cells(row, 2).Value) Like "x*" Then
                            str_ByMode = UCase(ws_def.Cells(row, 2).Value)
                            dict_Pmode2ByMode.Add performance_mode, str_ByMode
                            
                            If dict_ByMode2Index.Exists(str_ByMode) = False Then
                                dict_ByMode2Index.Add str_ByMode, cnt_ByMode
                                cnt_ByMode = cnt_ByMode + 1
                            End If
                        Else
                            enable_Parsing_Pmode2ByMode = False
                            Flag_Harvest_Pmode_Table_Parsed = False
                            TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & ",2), it is not the correct By_Mode for Parsing_Harvest_Pmode_Table. Error!!!"
                            TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & ",2), it is not the correct By_Mode for Parsing_Harvest_Pmode_Table. Error!!!"
                        End If
                    End If
                End If
            ElseIf enable_Parsing_FailFlag2DiableCore = True Then
                '''//FailFlag
                str_FailFlag = UCase(Trim(ws_def.Cells(row, 1).Value))
                
                '''//Harvest Core
                str_DiableCore = UCase(Trim(ws_def.Cells(row, 3).Value))
                
                '''//Check if str_Failflag exists in dict_FailFlag2DisableCore.
                If dict_FailFlag2DisableCore.Exists(str_FailFlag) = True Then
                    enable_Parsing_FailFlag2DiableCore = False
                    Flag_Harvest_Pmode_Table_Parsed = False
                    TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & ",1), HARV_Check:" & ws_def.Cells(row, 1).Value & " has the duplicate definition in dict_FailFlag2DisableCore for Parsing_Harvest_Pmode_Table. Error!!!"
                    TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & ",1), HARV_Check:" & ws_def.Cells(row, 1).Value & " has the duplicate definition in dict_FailFlag2DisableCore Parsing_Harvest_Pmode_Table. Error!!!"
                Else
                    dict_FailFlag2DisableCore.Add str_FailFlag, str_DiableCore
                    
                    '''//Add device_condition of the Failflag to the dictionary dict_FailFlagOfDisableCore2DevCondition.
                    '''//When Calculate_Harvest_Core_DSSC_Source, the FailFlag with device_condition is used to decide if Harvest Core should be disabled or not.
                    dict_FailFlagOfDisableCore2DevCondition.Add str_FailFlag, CBool(UCase(Trim(ws_def.Cells(row, 2).Value)))
                End If
                
                '''//Check if str_DiableCore exists in dict_DisableCore2FailFlag.
                If dict_DisableCore2FailFlag.Exists(str_DiableCore) = True Then
                    enable_Parsing_FailFlag2DiableCore = False
                    Flag_Harvest_Pmode_Table_Parsed = False
                    TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & ",3), DisableCore:" & ws_def.Cells(row, 3).Value & " has the duplicate definition in dict_DisableCore2FailFlag for Parsing_Harvest_Pmode_Table. Error!!!"
                    TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & ",3), DisableCore:" & ws_def.Cells(row, 3).Value & " has the duplicate definition in dict_DisableCore2FailFlag Parsing_Harvest_Pmode_Table. Error!!!"
                Else
                    dict_DisableCore2FailFlag.Add CLng(str_DiableCore), str_FailFlag
                End If
            ElseIf enable_Parsing_ByMode2CoreGroup = True Then
                If dict_columnOfTable.Count > 0 And dict_ByMode2Index.Count > 0 Then
                    '''init
                    cnt_overWriteSeq = 0
                    
                    '''Group1  Group2  Group3
                    str_ByMode = UCase(Trim(ws_def.Cells(row, dict_columnOfTable.Item(UCase("By_Mode"))).Value))
                    
                    '''//Check if ByMode exists in dict_ByMode2Index.
                    If dict_ByMode2Index.Exists(str_ByMode) = True Then
                        '''By_Mode
                        idx_ByMode = dict_ByMode2Index.Item(str_ByMode)
                    
                        If str_ByMode <> str_prv_ByMode Then
                            cnt_Condition_of_ByMode = 0
                        Else
                            cnt_Condition_of_ByMode = cnt_Condition_of_ByMode + 1
                        End If
                        
                        str_prv_ByMode = str_ByMode
                        ReDim Preserve HarvCoreByMode(idx_ByMode).condition(cnt_Condition_of_ByMode)
                        
                        '''Main_Core
                        HarvCoreByMode(idx_ByMode).condition(cnt_Condition_of_ByMode).MainCore = CLng(Trim(ws_def.Cells(row, dict_columnOfTable.Item(UCase("Main_Core"))).Value))
                        
                        '''Sequence[10:13]
                        HarvCoreByMode(idx_ByMode).condition(cnt_Condition_of_ByMode).bitStart_overWriteSeq = bit_overWriteSeq_start
                        HarvCoreByMode(idx_ByMode).condition(cnt_Condition_of_ByMode).bitStop_overWriteSeq = bit_overWriteSeq_stop
                        
                        '''//Check if column of Sequence contains "b####"
                        str_overWriteSeq = UCase(Trim(ws_def.Cells(row, dict_columnOfTable.Item(UCase("Sequence"))).Value))
                        
                        If str_overWriteSeq Like UCase("b*") Then '''ex:"B0000", "B1010"
                            str_overWriteSeq = Replace(str_overWriteSeq, "B", "")
                            
                            If Len(str_overWriteSeq) = (bit_overWriteSeq_stop - bit_overWriteSeq_start + 1) Then
                                For i = 1 To Len(str_overWriteSeq)
                                    ReDim Preserve lngAry_overWriteSeq(cnt_overWriteSeq)
                                    lngAry_overWriteSeq(cnt_overWriteSeq) = CLng(Mid(str_overWriteSeq, i, 1))
                                    cnt_overWriteSeq = cnt_overWriteSeq + 1
                                Next i
                                
                                HarvCoreByMode(idx_ByMode).condition(cnt_Condition_of_ByMode).overWriteSeq = lngAry_overWriteSeq
                            Else
                                enable_Parsing_ByMode2CoreGroup = False
                                Flag_Harvest_Pmode_Table_Parsed = False
                                TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & ",3), DisableCore:" & ws_def.Cells(row, 3).Value & " has the duplicate definition in dict_DisableCore2FailFlag for Parsing_Harvest_Pmode_Table. Error!!!"
                                TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & ",3), DisableCore:" & ws_def.Cells(row, 3).Value & " has the duplicate definition in dict_DisableCore2FailFlag Parsing_Harvest_Pmode_Table. Error!!!"
                                Exit Function
                            End If
                        End If
                        
                        '''FailFlag Name
                        HarvCoreByMode(idx_ByMode).condition(cnt_Condition_of_ByMode).Failflag = UCase(Trim(ws_def.Cells(row, dict_columnOfTable.Item(UCase("FailFlag Name"))).Value))
                        
                        '''Device Condition
                        If dict_FailFlag2DisableCore.Exists(HarvCoreByMode(idx_ByMode).condition(cnt_Condition_of_ByMode).Failflag) = True Then '''If not "N/A"
                            HarvCoreByMode(idx_ByMode).condition(cnt_Condition_of_ByMode).DevCondition = CBool(Trim(ws_def.Cells(row, dict_columnOfTable.Item(UCase("Device Condition"))).Value))
                        End If
                        
                        '''GroupName-loop by strAry_HarvestCoreGroupName
                        HarvCoreByMode(idx_ByMode).condition(cnt_Condition_of_ByMode).dict_GroupName2CoreGroup.RemoveAll
                        For idx_GroupName = 0 To UBound(strAry_HarvestCoreGroupName)
                            str_GroupName = UCase(strAry_HarvestCoreGroupName(idx_GroupName))
                            str_Temp = ws_def.Cells(row, dict_columnOfTable.Item(str_GroupName)).Value
                            
                            If str_Temp <> "" Then
                                split_content = Split(str_Temp, ",")
                                
                                For i = 0 To UBound(split_content)
                                    ReDim Preserve lngAry_coreGroup(i)
                                    lngAry_coreGroup(i) = CLng(split_content(i))
                                Next i
                                
                                HarvCoreByMode(idx_ByMode).condition(cnt_Condition_of_ByMode).dict_GroupName2CoreGroup.Add str_GroupName, lngAry_coreGroup
                            End If
                        Next idx_GroupName
                    Else
                        enable_Parsing_ByMode2CoreGroup = False
                        Flag_Harvest_Pmode_Table_Parsed = False
                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & ",1), By_Mode:" & str_ByMode & ", it is undefined By_Mode for Parsing_Harvest_Pmode_Table. Error!!!"
                        TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & ",1), By_Mode:" & str_ByMode & ", it is undefined By_Mode for Parsing_Harvest_Pmode_Table. Error!!!"
                    End If
                End If
            ElseIf enable_Parsing_EnableCore2Fstp = True Then
                str_EnableCore = UCase(Trim(ws_def.Cells(row, 1).Value))
                str_Fstp = UCase(Trim(ws_def.Cells(row, 2).Value))
                
                '''//Check if str_EnableCore exists in dict_EnableCore2Fstp.
                If dict_EnableCore2Fstp.Exists(CLng(str_EnableCore)) = True Then
                    enable_Parsing_EnableCore2Fstp = False
                    Flag_Harvest_Pmode_Table_Parsed = False
                    TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & ",1), EnableCore:" & ws_def.Cells(row, 1).Value & " has the duplicate definition in dict_EnableCore2Fstp for Parsing_Harvest_Pmode_Table. Error!!!"
                    TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & ",1), EnableCore:" & ws_def.Cells(row, 1).Value & " has the duplicate definition in dict_EnableCore2Fstp Parsing_Harvest_Pmode_Table. Error!!!"
                Else
                    dict_EnableCore2Fstp.Add CLng(str_EnableCore), str_Fstp
                    ReDim Preserve strAry_HarvestCoreFstpName(cnt_Fstp)
                    strAry_HarvestCoreFstpName(cnt_Fstp) = str_Fstp
                    cnt_Fstp = cnt_Fstp + 1
                End If
                
                '''//Check if str_Fstp exists in dict_Fstp2EnableCore.
                If enable_Parsing_EnableCore2Fstp = True Then
                    If dict_Fstp2EnableCore.Exists(str_Fstp) = True Then
                        enable_Parsing_EnableCore2Fstp = False
                        Flag_Harvest_Pmode_Table_Parsed = False
                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & ",2), FSTP:" & ws_def.Cells(row, 2).Value & " has the duplicate definition in dict_DisableCore2FailFlag for Parsing_Harvest_Pmode_Table. Error!!!"
                        TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & ",2), FSTP:" & ws_def.Cells(row, 2).Value & " has the duplicate definition in dict_DisableCore2FailFlag Parsing_Harvest_Pmode_Table. Error!!!"
                    Else
                        dict_Fstp2EnableCore.Add str_Fstp, CLng(str_EnableCore)
                    End If
                End If
            End If
        Next row
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Parsing_Harvest_Pmode_Table"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Parsing_Harvest_Pmode_Table"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210510: Modified to parse HARV_Pmode_Table and HARVMappingTable for Harvest Core DSSC.
Public Function Parsing_Harvest_Mapping_Table(sheetName As String)
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim isSheetFound As Boolean
    Dim row As Long, col As Long
    Dim i As Long
    Dim enableRowParsing As Boolean
    Dim strAry_Fstp() As String
    Dim cnt_Fstp As Integer
    Dim dict_columnOfTable As New Dictionary
    Dim str_Fstp As String
    '''columns
    Dim col_Sequence As Long
    Dim cnt_Sequence As Long
    '''variables
    Dim str_PatternName As String
    Dim idx_PatternGrp As Long
    '''variables
    Dim idx_Fstp As Long
    Dim strAry_Fstp_BitSequence() As String
    Dim col_Fstp As Long
    Dim split_content() As String
    Dim idx_BitSequence As Long
    Dim lngAry_BitSequence() As Long
On Error GoTo errHandler
    '''*****************************************************************'''
    '''//Check if the sheet exists
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        enableRowParsing = False
        cnt_Fstp = 0
        dict_columnOfTable.RemoveAll
        cnt_Sequence = 0
        idx_PatternGrp = -1
        str_PatternName = ""
        
        For row = 1 To MaxRow
            '''//Find the header.
            If LCase(ws_def.Cells(row, 1).Value) Like LCase("PatternName") And LCase(ws_def.Cells(row, 2).Value) Like LCase("Sequence") Then
                '''//initialize patternName
                str_PatternName = ""
                
                '''//Check each column in row of the header.
                For col = 1 To maxcol
                    If UCase(Trim(ws_def.Cells(row, col).Value)) <> "" Then
                        str_Fstp = ""
                        
                        If dict_columnOfTable.Exists(UCase(Trim(ws_def.Cells(row, col).Value))) = True Then
                            If dict_columnOfTable.Item(UCase(Trim(ws_def.Cells(row, col).Value))) = col Then
                                '''Do nothing
                            Else
                                enableRowParsing = False
                                TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & "," & col & "), columns of the header are inconsistent for Parsing_Harvest_Mapping_Table. Error!!!"
                                TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & "," & col & "), columns of the header are inconsistent for Parsing_Harvest_Mapping_Table. Error!!!"
                            End If
                        Else
                            dict_columnOfTable.Add UCase(Trim(ws_def.Cells(row, col).Value)), col
                        End If
                        
                        If LCase(Trim(ws_def.Cells(row, col).Value)) Like "fstp*" Then
                            str_Fstp = UCase(Trim(ws_def.Cells(row, col).Value))
                            
                            '''//Check if the header of FSTP column is defined in dict_Fstp2EnableCore of HARV_Pmode_Table.
                            If dict_Fstp2EnableCore.Exists(str_Fstp) Then
                                ReDim Preserve strAry_Fstp(cnt_Fstp)
                                strAry_Fstp(cnt_Fstp) = str_Fstp
                                cnt_Fstp = cnt_Fstp + 1
                                enableRowParsing = True
                                cnt_Sequence = 0
                            Else
                                enableRowParsing = False
                                TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & "," & col & "), FSTP:" & ws_def.Cells(row, col).Value & " isn't defined in HARV_Pmode_Table for Parsing_Harvest_Mapping_Table. Error!!!"
                                TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & "," & col & "), FSTP:" & ws_def.Cells(row, col).Value & " isn't defined in HARV_Pmode_Table for Parsing_Harvest_Mapping_Table. Error!!!"
                                Exit For
                            End If
                        End If
                    End If
                Next col
                
                If enableRowParsing = True Then
                    If dict_columnOfTable.Exists(UCase("PatternName")) And dict_columnOfTable.Exists(UCase("Sequence")) Then
                        row = row + 1
                        idx_PatternGrp = idx_PatternGrp + 1
                        ReDim Preserve HarvCoreDSSC_BitSequence(idx_PatternGrp)
                    Else
                        enableRowParsing = False
                        Exit For
                    End If
                    
                    '''//Check if all FSTP are shown in the header.
                    For idx_Fstp = 0 To UBound(strAry_HarvestCoreFstpName)
                        If dict_columnOfTable.Exists(strAry_HarvestCoreFstpName(idx_Fstp)) = True Then
                            ReDim strAry_Fstp_BitSequence(dict_Fstp2EnableCore.Count - 1) As String
                        Else
                            enableRowParsing = False
                            TheExec.Datalog.WriteComment "sheet:" & sheetName & ", it doesn't have " & strAry_HarvestCoreFstpName(idx_Fstp) & " in the header for Parsing_Harvest_Mapping_Table. Error!!!"
                            TheExec.ErrorLogMessage "sheet:" & sheetName & ", it doesn't have " & strAry_HarvestCoreFstpName(idx_Fstp) & " in the header for Parsing_Harvest_Mapping_Table. Error!!!"
                            Exit For
                        End If
                    Next idx_Fstp
                End If
            End If
            
            If enableRowParsing = True Then
                col_Sequence = dict_columnOfTable.Item(UCase("Sequence"))
                
                If ws_def.Cells(row, col_Sequence).Value <> "" Then
                    '''//Check if bit sequence is correct.
                    If CLng(ws_def.Cells(row, col_Sequence).Value) = cnt_Sequence Then
                        Flag_Harvest_Mapping_Table_Parsed = True
                        
                        '''PatternName
                        If ws_def.Cells(row, dict_columnOfTable.Item(UCase("PatternName"))).Value <> "" Then
                            If str_PatternName <> "" Then
                                str_PatternName = str_PatternName & "," & ws_def.Cells(row, dict_columnOfTable.Item(UCase("PatternName"))).Value
                            Else
                                str_PatternName = ws_def.Cells(row, dict_columnOfTable.Item(UCase("PatternName"))).Value
                            End If
                        End If
                        
                        If str_PatternName <> "" Then
                            HarvCoreDSSC_BitSequence(idx_PatternGrp).Pattern = Split(str_PatternName, ",")
                        End If
                        
                        '''Sequence
                        '''Harvest Result0
                        '''FSTP0   FSTP1   FSTP2   FSTP3   FSTP4   FSTP5   FSTP6   FSTP7   FSTP8   FSTP9
                        '''//Get bitSequence of each FSTP
                        HarvCoreDSSC_BitSequence(idx_PatternGrp).bitSeq_Core.RemoveAll
                        
                        For idx_Fstp = 0 To UBound(strAry_HarvestCoreFstpName)
                            col_Fstp = dict_columnOfTable.Item(strAry_HarvestCoreFstpName(idx_Fstp))
                            
                            '''dict_Fstp2EnableCore
                            If strAry_Fstp_BitSequence(idx_Fstp) <> "" Then
                                strAry_Fstp_BitSequence(idx_Fstp) = strAry_Fstp_BitSequence(idx_Fstp) & "," & ws_def.Cells(row, col_Fstp).Value
                            Else
                                strAry_Fstp_BitSequence(idx_Fstp) = ws_def.Cells(row, col_Fstp).Value
                            End If
                            
                            split_content = Split(strAry_Fstp_BitSequence(idx_Fstp), ",")
                            
                            '''//Add the bitSequence of each FTSP into HarvCoreDSSC_BitSequence(idx_PatternGrp)
                            ReDim lngAry_BitSequence(UBound(split_content)) As Long
                            For idx_BitSequence = 0 To UBound(split_content)
                                lngAry_BitSequence(idx_BitSequence) = CLng(split_content(idx_BitSequence))
                            Next idx_BitSequence
                            HarvCoreDSSC_BitSequence(idx_PatternGrp).bitSeq_Core.Add strAry_HarvestCoreFstpName(idx_Fstp), lngAry_BitSequence
                        Next idx_Fstp
                    Else
                        enableRowParsing = False
                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell(" & row & "," & col & "), Sequence has the incorrect bit. Error!!!"
                        TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell(" & row & "," & col & "), Sequence has the incorrect bit. Error!!!"
                        Exit For
                    End If
                    
                    Flag_Harvest_Mapping_Table_Parsed = True
                    cnt_Sequence = cnt_Sequence + 1
                    HarvCoreDSSC_BitSequence(idx_PatternGrp).cnt_BitSequence = cnt_Sequence
                End If
            End If
        Next row
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Parsing_Harvest_Mapping_Table"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Parsing_Harvest_Mapping_Table"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210610: Modified to check Flag_Harvest_Core_DSSC_Ready for Calculate_Harvest_Core_DSSC_Source.
'20210604: Modified to check the pattern for Calculate_Harvest_Core_DSSC_Source first, then check Pattern_Pmode/By_Mode/Harvest_Core_Group.
'20210528: Modified to add the optional argument Optional inst_Pmode_keyword As String = "" and Optional inst_By_Mode As String = "".
'20210517: Modified to check instance keywords, ex: "MGX001".
'20210512: Modified to parse HARV_Pmode_Table and HARVMappingTable for Harvest Core DSSC.
Public Function Calculate_Harvest_Core_DSSC_Source(instance As String, performance_mode As String, Pattern As String, DigSrc_pin As PinList, DigSrc_SignalName As String, Optional inst_Pmode_keyword As String = "", Optional inst_By_Mode As String = "")
    Dim site As Variant
    Dim split_content() As String
    Dim str_instance_pmode As String
    Dim str_ByMode As String
    Dim idx_ByMode As Long
    Dim pos_ByMode As Long
    Dim i As Long
    Dim str_HarvestCoreGroup As String
    Dim selsram_expand_array() As Long
    Dim gotHarvestCoreGroupName As Boolean
    Dim idx As Long
    Dim idx_BitSequenceGroup As Long
    Dim idx_pattern As Long
    Dim idx_flag_condition As Long
    Dim lngAry_temp_BitSequence() As Long
    Dim cnt_temp_BitSequence As Long
    Dim lngAry_selected_Core() As Long
    Dim cnt_selected_Core As Long
    Dim lngAry_temp_Core() As Long
    Dim str_FailFlag As String
    Dim lngAry_single_BitSequence() As Long
    Dim str_Fstp As String
    Dim str_singleSite_FSTP_info As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Warning!!! please check the flow table to prevent Harvesting from 2 Cores failed.
'''//==================================================================================================================================================================================//'''
    '''20210610: Modified to check Flag_Harvest_Core_DSSC_Ready for Calculate_Harvest_Core_DSSC_Source.
    If Flag_Harvest_Core_DSSC_Ready = False Then
        Exit Function
    End If

    '''init
    pos_ByMode = -1
    str_HarvestCoreGroup = ""
    gotHarvestCoreGroupName = False
    idx_BitSequenceGroup = -1
    idx_flag_condition = -1
    cnt_selected_Core = 0
    str_Fstp = ""
    str_singleSite_FSTP_info = ""
    
    '''//Check if the matched pattern can be found in HarvMappingTable.
    '''//Find the matched pattern from HarvMappingTable, then get idx_BitSequenceGroup to find the bit sequence for each Harvest Core.
    '''20210604: Modified to check the pattern for Calculate_Harvest_Core_DSSC_Source first, then check Pattern_Pmode/By_Mode/Harvest_Core_Group.
    For idx = 0 To UBound(HarvCoreDSSC_BitSequence)
        For idx_pattern = 0 To UBound(HarvCoreDSSC_BitSequence(idx).Pattern)
            If UCase(Pattern) Like UCase("*" & HarvCoreDSSC_BitSequence(idx).Pattern(idx_pattern) & "*") Then
                idx_BitSequenceGroup = idx
                Exit For
            End If
        Next idx_pattern
        
        If idx_BitSequenceGroup <> -1 Then
            Exit For
        End If
    Next idx

    '''//If no matched pattern in HarvMappingTable, exit function.
    If idx_BitSequenceGroup = -1 Then
        Exit Function
    End If
    
    '''//Split the instance name by "_" and find the position of ByMode.
    split_content = Split(instance, "_")
    For i = 0 To UBound(split_content)
        If i + 2 < UBound(split_content) Then
            '''***********************************************************************************************************************************************'''
            '''//Get instance_pmode, ByMode, and HarvestCoreGroup, ex: "MGX001_X4_GROUP1".
            '''For example, the instance name "MG001_GpuTd_MGX001_X4_Group1_PP_STNA0_L_PL00_PL01_PL02_PL03_PL04_SC_CFXX_TDF_COM_AUT_MGXXXX_DM_MLTFSTP_XOR_BV".
            '''"MGX001" is the pattern_pmode, "X4" is the By_Mode, and "Group1" is the Harvest core group.
            '''***********************************************************************************************************************************************'''
            '''//Check if pattern_pmode exists in the instance name, ex: "MGX001".
            '''20210517: Modified to check instance keywords, ex: "MGX001".
            '''20210528: Modified to add the optional argument Optional inst_Pmode_keyword As String = "" and Optional inst_By_Mode As String = "".
            If dict_Pmode2ByMode.Exists(UCase(split_content(i))) = True Then
                '''//If pattern_pmode exists, Get pattern_pmode from the instance name.
                str_instance_pmode = UCase(split_content(i))
                inst_Pmode_keyword = str_instance_pmode
                
                '''//Check if By_Mode from the instance name matches By_Mode of pattern_pmode in dict_Pmode2ByMode(from HarvMappingTable).
                If UCase(split_content(i + 1)) = UCase(dict_Pmode2ByMode.Item(UCase(str_instance_pmode))) Then
                    str_ByMode = dict_Pmode2ByMode.Item(UCase(str_instance_pmode))
                    inst_By_Mode = str_ByMode
                    
                    '''//If that, get HarvestCoreGroup from position idx_ByMode+1 in the instance name, ex: "Group1".
                    idx_ByMode = dict_ByMode2Index.Item(UCase(str_ByMode))
                    pos_ByMode = i + 1
                    
                    '''//If no matched By_Mode, exit function.
                    If pos_ByMode = UBound(split_content) Then
                        Exit Function
                    Else
                        '''//Get the Harvest Core group(HarvestCoreGroup), ex: "GROUP1".
                        str_HarvestCoreGroup = UCase(split_content(pos_ByMode + 1))
                    End If
                    
                    Exit For
                End If
            End If
        Else
            Exit Function
        End If
    Next i
    
    '''//Check if str_HarvestCoreGroup exists in strAry_HarvestCoreGroupName.
    '''//If HarvestCoreGroup doesn't exist in the instance name, exit function.
    If str_HarvestCoreGroup = "" Then
        Exit Function
    Else
        '''//Check if HarvestCoreGroup exists in "HARV_Pmode_Table", ex: "Group1", "Group2", "Group3"
        If dict_HarvestCoreGroup2Index.Exists(str_HarvestCoreGroup) = True Then
            gotHarvestCoreGroupName = True
        Else
            Exit Function
        End If
    End If
        
    For Each site In TheExec.sites
        '''init
        cnt_selected_Core = 0
        str_FailFlag = ""
        str_Fstp = ""
        idx_flag_condition = -1
        str_singleSite_FSTP_info = "Site:" & site
        
        
        '''//If HarvestCoreGroup exists, check FailFlag with device_condition of the matched MainCore.
        For idx = 0 To UBound(HarvCoreByMode(idx_ByMode).condition)
            '''//Check if the main core has the correct failFlag.
            If dict_FailFlag2DisableCore.Exists(UCase(HarvCoreByMode(idx_ByMode).condition(idx).Failflag)) = True Then '''ex: "F_gfx_Core0".
                str_FailFlag = UCase(HarvCoreByMode(idx_ByMode).condition(idx).Failflag)
                    
                '''//If flagstate of FailFlag matches device_condition of the main core, get idx_flag_condition to find core group for Harvest Core group.
                If CBool(TheExec.sites.Item(site).FlagState(str_FailFlag)) = HarvCoreByMode(idx_ByMode).condition(idx).DevCondition Then
                    idx_flag_condition = idx
                    Exit For
                End If
            Else '''ex: "N/A".
                idx_flag_condition = idx
            End If
        Next idx
        
        '''If condition of By_Mode is found, start to get the selected Core Group for Harvest Core DSSC.
        If idx_flag_condition <> -1 Then
            cnt_temp_BitSequence = HarvCoreDSSC_BitSequence(idx_BitSequenceGroup).cnt_BitSequence
            ReDim lngAry_temp_BitSequence(cnt_temp_BitSequence - 1) As Long
        
            '''//Check flagstate of Failflag for each Core in HarvestCoreGroup, and decide the final CoreGrp.
            If HarvCoreByMode(idx_ByMode).condition(idx_flag_condition).dict_GroupName2CoreGroup.Exists(str_HarvestCoreGroup) Then
                '''init
                str_FailFlag = ""
                cnt_selected_Core = 0
                
                '''//Check flagstate each core in CoreGrp, and update the core array.
                lngAry_temp_Core = HarvCoreByMode(idx_ByMode).condition(idx_flag_condition).dict_GroupName2CoreGroup.Item(str_HarvestCoreGroup)
                For idx = 0 To UBound(lngAry_temp_Core)
                    str_FailFlag = UCase(dict_DisableCore2FailFlag.Item(lngAry_temp_Core(idx)))
                        
                    '''//Diable the failed Harvest Core.
                    If CBool(TheExec.sites.Item(site).FlagState(str_FailFlag)) = dict_FailFlagOfDisableCore2DevCondition.Item(str_FailFlag) Then
                        '''If failFlag=True, skip the failed core.
                    Else
                        '''//For the Harvest Core not failed, add it to array of the Selected_Core.
                        ReDim Preserve lngAry_selected_Core(cnt_selected_Core) As Long
                        lngAry_selected_Core(cnt_selected_Core) = lngAry_temp_Core(idx)
                        cnt_selected_Core = cnt_selected_Core + 1
                    End If
                Next idx
                
                '''//If array of the Selected_Core is not empty, it means that Harvest Core is available...
                If cnt_selected_Core > 0 Then
                    '''//Assemble bitArray for DSSC DigSrc.
                    For idx = 0 To UBound(lngAry_selected_Core)
                        str_Fstp = UCase(dict_EnableCore2Fstp.Item(lngAry_selected_Core(idx)))
                        lngAry_single_BitSequence = HarvCoreDSSC_BitSequence(idx_BitSequenceGroup).bitSeq_Core.Item(str_Fstp)
                        
                        '''//Do bit-wise "OR" calculation of bit of each Core for Harvest Core DSSC.
                        For i = 0 To UBound(lngAry_temp_BitSequence)
                            lngAry_temp_BitSequence(i) = lngAry_temp_BitSequence(i) Or lngAry_single_BitSequence(i)
                        Next i
                    Next idx
                    
                    '''//Overwrite BitSequence[10:13].
                    '''//Sequence[10:13] from HarvMappingTable will be used to replace bit10~13 in Harvest Core DSSC bit sequence.
                    If HarvCoreByMode(idx_ByMode).condition(idx_flag_condition).bitStart_overWriteSeq > -1 Then
                        For i = 0 To UBound(HarvCoreByMode(idx_ByMode).condition(idx_flag_condition).overWriteSeq)
                            lngAry_temp_BitSequence(HarvCoreByMode(idx_ByMode).condition(idx_flag_condition).bitStart_overWriteSeq + i) = HarvCoreByMode(idx_ByMode).condition(idx_flag_condition).overWriteSeq(i)
                        Next i
                    End If
                    
                    '''//Print the flagstate of failFlag for each Harvest core.
                    str_singleSite_FSTP_info = str_singleSite_FSTP_info & "," & UCase(HarvCoreByMode(idx_ByMode).condition(idx_flag_condition).Failflag) & "=" & CStr(CBool(TheExec.sites.Item(site).FlagState(UCase(HarvCoreByMode(idx_ByMode).condition(idx_flag_condition).Failflag))))
                    
                    '''//Print the bit sequence of Harvest Core DSSC.
                    str_singleSite_FSTP_info = str_singleSite_FSTP_info & ", Src Bits = " & cnt_temp_BitSequence & ", HarvestSourceCode [ First(L) ==> Last(R) ] :"
                    
                    For i = 0 To UBound(lngAry_temp_BitSequence)
                        str_singleSite_FSTP_info = str_singleSite_FSTP_info & CStr(lngAry_temp_BitSequence(i))
                    Next i
                    
                    '''//Print FSTP Harvest core DSSC info
                    TheExec.Datalog.WriteComment str_singleSite_FSTP_info
                    
                    '''//Setup DSSC DigSrc with Harvest Core DSSC pattern and Harv_core_DSSC_array.
                    Call DSSC_SetupDigSrcWave(Pattern, DigSrc_pin, DigSrc_SignalName, cnt_temp_BitSequence, lngAry_temp_BitSequence)
                Else
                    '''***********************************************************************************************************'''
                    '''Warning!!! please check the flow table to prevent Harvesting from 2 Cores failed.
                    '''***********************************************************************************************************'''
                    '''ToDo: Maybe we can use siteMask to mask the failed site...
                    TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                    TheExec.Datalog.WriteComment "Site:" & site & "," & performance_mode & ",instance:" & instance & ", no core passed for Calculate_Harvest_Core_DSSC_Source. Error!!!"
                    TheExec.ErrorLogMessage "Site:" & site & "," & performance_mode & ",instance:" & instance & ", no core passed for Calculate_Harvest_Core_DSSC_Source. Error!!!"
                End If
            Else
                TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                TheExec.Datalog.WriteComment "Site:" & site & "," & performance_mode & ",instance:" & instance & ", it has the incorrect Group:" & str_HarvestCoreGroup & " for Calculate_Harvest_Core_DSSC_Source. Error!!!"
                TheExec.ErrorLogMessage "Site:" & site & "," & performance_mode & ",instance:" & instance & ", it has the incorrect Group:" & str_HarvestCoreGroup & " for Calculate_Harvest_Core_DSSC_Source. Error!!!"
            End If
        Else
            TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
            TheExec.Datalog.WriteComment "Site:" & site & "," & performance_mode & ",instance:" & instance & ", it doesn't have any matached device condition for Calculate_Harvest_Core_DSSC_Source. Error!!!"
            TheExec.ErrorLogMessage "Site:" & site & "," & performance_mode & ",instance:" & instance & ", it doesn't have any matached device condition for Calculate_Harvest_Core_DSSC_Source. Error!!!"
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Calculate_Harvest_Core_DSSC_Source"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Calculate_Harvest_Core_DSSC_Source"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210721: Modified to use gb_IDS_hi_limit(VddBinStr2Enum(powerDomain), passBinCut) for OtherRail.
'20210720: Modified to check if Efuse category "Product_Identifier" exists in Efuse_BitDef_Table.
'20210720: Modified to define ids_for_offline as localVariable to store IDS values from offline simulation.
'20210716: Modified to decide Efuse setReadDecimal or setWriteDecimal for generate_offline_IDS_IGSim_Parallel.
'20210715: Modified to write offline simulation values into field.FuseMeasureValue for Efuse product voltages.
'20210715: Modified to align IDS offline simulation with unit: mA.
'20210714: Modified to generated offline Efuse product voltage.
'20210712: Modified to generate random values for BinCut PassBin.
'20210709: Modified to move the vbt function generate_offline_IDS_IGSim_Parallel from the vbt module VBT_LIB_VDD_Binning to LIB_VDD_BINNING_customer.
'20210709: Modified to move the vbt code for Efuse product voltage offline simulation from the vbt function Read_DVFM_To_GradeVDD to generate_offline_IDS_IGSim_Parallel.
'20210709: Modified to move the vbt code for Product_Identifier offline simulation from the vbt function Read_DVFM_To_GradeVDD to generate_offline_IDS_IGSim_Parallel.
'20210708: Discussed the vbt code of updating offline simulation IDS values with ZQLIN, he suggested us to keep the vbt code here for those powerPins with the correct Efuse category.
'20210708: Modified to merge the branches of the vbt function generate_offline_IDS_IGSim_Parallel.
'20210707: Modified to check if it can use powerDomain or 1st powerPin in PinGroup.
'20210707: Modified to find Efuse Category for the BinCut performance mode by checking the current BinCut testjob.
'20210707: As per discussion with TSMC SWLINZA, BinCut retest and correlation should use Efuse IDS values only, not Efuse Product_Identifier!!!
'20210617: Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
'20201228: Patty asked us to adapt Efuse object vbt code.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20201112: Modified to use the dictionary "dict_IsCorePowerInBinCutFlowSheet".
'20200730: Modified to add the EnableWord "VDDBinning_Offline_AllPattPass" for Offline simulation with all patterns pass.
'20200130: Modified to check powerDomain and powerPin.
'20190715: Modified to get IDS hi_limit by using the new power_seq.
Public Function generate_offline_IDS_IGSim_Parallel() As Long
    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim siteLng_start_PassBin As New SiteLong
    Dim siteLng_final_PassBin As New SiteLong
    Dim p_mode As Integer
    Dim strAry_powerSeq() As String
    Dim powerPin As String
    Dim str_IDS_PowerDomain As String
    Dim highest_Performance_mode As String
    Dim upperbound As Variant, lowerbound As Variant, tempbound As Variant
    Dim dbl_generated_BV As Double
    Dim dbl_generated_Product As Double
    Dim str_Efuse_read_ProductIdentifier As String
    Dim str_Efuse_read_pmode As String
    Dim idx_step As Long
    Dim got_Efuse_category As Boolean
    Dim dbl_quotient As Double
    Dim dbl_Efuse_round_number As Double
    Dim ids_for_offline(1 To MaxBincutPowerdomainCount) As New SiteDouble
    Dim dbl_ratio_Ids As Double
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
'''2. As per discussion with TSMC SWLINZA, BinCut retest and correlation should use Efuse IDS values only, not Efuse Product_Identifier!!!
'''3. As per discussion with TSMC SWLINZA, for powerPin group, it should use 1st powerPin to check IDS limit of powerPin group, 20210707.
'''ex: powerGroup: VDD_FIXED_GRP, and its 1st powerPin: VDD_FIXED, so that compare IDS value of VDD_FIXED with IDS_limit of VDD_FIXED_GRP. It must have Efuse category in Efuse_BitDef_Table to store IDS for VDD_FIXED.
'''4. Since Efuse obj vbt code only provides one time permission to set the item, it can't update Efuse IDS in the vbt function generate_offline_IDS_IGSim_Parallel...
'''//==================================================================================================================================================================================//'''
    '''//Get Efuse catergory of "Product_Identifier" to "read" PassBin (Product_Identifier+1).
    '''20210720: Modified to check if Efuse category "Product_Identifier" exists in Efuse_BitDef_Table.
    If dict_EfuseCategory2BinCutTestJob.Exists(UCase("Product_Identifier")) = True Then
        str_Efuse_read_ProductIdentifier = get_Efuse_category_by_BinCut_testJob("read", "Product_Identifier")
    Else
        TheExec.Datalog.WriteComment "Efuse category:" & "Product_Identifier" & ", it doesn't exists in Efuse_BitDef_Table. Please check Efuse_BitDef_Table for generate_offline_IDS_IGSim_Parallel. Error!!!"
        TheExec.ErrorLogMessage "Efuse category:" & "Product_Identifier" & ", it doesn't exists in Efuse_BitDef_Table. Please check Efuse_BitDef_Table for generate_offline_IDS_IGSim_Parallel. Error!!!"
    End If
    
    '''//Check if it has to do simulation of all sites Bin1 for Opensocket or offline.
    For Each site In TheExec.sites
        '''//Generate BinCut PassBin for opensocket/offline simulation.
        '''20210712: Modified to generate random values for BinCut PassBin.
        If EnableWord_VDDBinning_Offline_AllPattPass = True Or EnableWord_Golden_Default = True Then '''all sites are simulated as Bin1.
            siteLng_start_PassBin(site) = 1
            siteLng_final_PassBin(site) = 1
        Else '''Generate random PassBin number for opensocket/offline simulation.
            siteLng_start_PassBin(site) = (Rnd * Total_Bincut_Num) / 1
            siteLng_final_PassBin(site) = (Rnd * Total_Bincut_Num) / 1
            
            If siteLng_start_PassBin(site) = 0 Then
                siteLng_start_PassBin(site) = 1
            End If
            
            If siteLng_start_PassBin(site) > siteLng_final_PassBin(site) Then
                siteLng_final_PassBin(site) = siteLng_start_PassBin(site)
            End If
        End If
    Next site
    
    '''//Update BinCut PassBin to Efuse Product_Identifier for non-CP1.
    If str_Efuse_read_ProductIdentifier <> "" Then
        '''For project with Efuse DSP vbt code.
        CFGFuse.Category(CFGIndex(str_Efuse_read_ProductIdentifier)).Read.Value = siteLng_final_PassBin.Subtract(1)
    Else
        TheExec.Datalog.WriteComment "Efuse category:" & "Product_Identifier" & ", it wasn't fused before the currentJob, so that skip generating offline product_identifier."
    End If
    
    For i = 0 To UBound(pinGroup_BinCut)
        '''//Get 1st powerPin of the powerDomain.
        '''20210707: As per discussion with TSMC SWLINZA, for powerPin group, it should use 1st powerPin to check IDS limit of powerPin group.
        '''ex: powerGroup: VDD_FIXED_GRP, and its 1st powerPin: VDD_FIXED, so that compare IDS value of VDD_FIXED with IDS_limit of VDD_FIXED_GRP.
        '''It must have Efuse category in Efuse_BitDef_Table to store IDS for VDD_FIXED.
        If domain2pinDict.Exists(UCase(pinGroup_BinCut(i))) Then
            powerPin = Get1stPinFromPingroup(VddbinDomain2Pin(UCase(pinGroup_BinCut(i)))) '''get 1st powerPin from powerDomain.
        ElseIf pin2domainDict.Exists(UCase(pinGroup_BinCut(i))) Then
            powerPin = pinGroup_BinCut(i)
        Else
            powerPin = ""
            TheExec.Datalog.WriteComment "site:" & site & ",powerDomain:" & powerPin & ",Efuse category:" & ",it can't get Efuse category to get Efuse resolution for generate_offline_IDS_IGSim_Parallel. Please check Efuse_BitDef_Table or instances of Harvest result update. Error!!!"
            TheExec.ErrorLogMessage "site:" & site & ",powerDomain:" & powerPin & ",Efuse category:" & ",it can't use Efuse category to get Efuse resoultion for generate_offline_IDS_IGSim_Parallel. Please check Efuse_BitDef_Table or instances of Harvest result update. Error!!!"
        End If
        
        If powerPin <> "" Then
            '''init
            str_IDS_PowerDomain = IDS_for_BinCut(VddBinStr2Enum(powerPin)).ids_name(site)
            
            '''//IDS_for_offline is the new siteDouble defined in "LIB_Vdd_Binning_GlobalVariable".
            ids_for_offline(VddBinStr2Enum(powerPin))(site) = 0
            
            '''//For powerDomain with Harvesting, it might get no IDS name after parsing Efuse_BitDef_Table.
            '''So that manully add the keyword about Harvest core number to the IDS name for the powerDomain.
            '''ToDo: Please check Harvest core number in Efuse_BitDef_Table and discuss this with project owner...
            '''20210708: Modified to merge the branches of the vbt function generate_offline_IDS_IGSim_Parallel.
            If str_IDS_PowerDomain = "" Then
                str_IDS_PowerDomain = "IDS_" & UCase(powerPin) & "_10" '''Please check the harvest core number...
                IDS_for_BinCut(VddBinStr2Enum(powerPin)).ids_name(site) = str_IDS_PowerDomain
            End If
            
            '''//Check if str_IDS_PowerDomain is the correct Efuse category in Efuse_BitDef_Table.
            '''20210707: Modified to find Efuse Category for the BinCut performance mode by checking the current BinCut testjob.
            If dict_EfuseCategory2BinCutTestJob.Exists(UCase(str_IDS_PowerDomain)) = True Then
                got_Efuse_category = True
            Else
                got_Efuse_category = False
            End If
            
            '''//Check if powerPin is BinCut CorePower or OtherRail by using the dictionary "dict_IsCorePowerInBinCutFlowSheet".
            If dict_IsCorePowerInBinCutFlowSheet.Exists(powerPin) = True Then
                If dict_IsCorePowerInBinCutFlowSheet.Item(powerPin) = True Then '''CorePower
                    '''//Get the highest performance mode from power_Seq of the PowerDomain.
                    '''We use the highest performance mode to get its CP IDS_limit.
                    highest_Performance_mode = BinCut_Power_Seq(VddBinStr2Enum(powerPin)).Power_Seq(UBound(BinCut_Power_Seq(VddBinStr2Enum(powerPin)).Power_Seq))
                    
                    For Each site In TheExec.sites
                        '''//Use IDS limit as the upper bound of random number.
                        '''//Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
                        upperbound = BinCut(VddBinStr2Enum(highest_Performance_mode), siteLng_final_PassBin(site)).IDS_CP_LIMIT(0) '''unit: mA.
                        
                        '''//Get lowerbound.
                        '''CorePower has the Efuse category about IDS name, and Efuse always rounds up DC_TEST IDS values as 1# resolution at least.
                        '''Use Efuse resolution as lowerbound for random numbers.
                        '''20210715: Modified to align IDS offline simulation with unit: mA.
                        If got_Efuse_category = True Then
                            '''For project with Efuse DSP vbt code.
                            lowerbound = 1# * CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion '''unit: mA
                        Else
                            lowerbound = 0
                        End If
                    
                        '''//Check if lowerbound and upperbound are correct.
                        If lowerbound > upperbound Then
                            tempbound = lowerbound
                            lowerbound = upperbound
                            upperbound = tempbound
                        End If
                        
                        '''//Ratio for IDS offline simulation.
                        If str_Efuse_read_ProductIdentifier <> "" Then
                            dbl_ratio_Ids = 0.7
                        Else
                            dbl_ratio_Ids = 0.85
                        End If
                        
                        '''//Generate the random value for IDS.
                        ids_for_offline(VddBinStr2Enum(powerPin))(site) = UserSiteVar(str_IDS_PowerDomain, upperbound, lowerbound) * dbl_ratio_Ids '''unit: mA
                        
                        '''//If Efuse category of IDS name exists, use the Efuse method to process offline IDS value.
                        If got_Efuse_category = True Then
                            '''//Update offline IDS values to Efuse data structure if str_IDS_PowerDomain is the correct Efuse category.
                            '''For project with Efuse DSP vbt code.
                            '''//For Efuse IDS category, it should round up DC_TEST IDS values with Efuse resolution.
                            dbl_quotient = ids_for_offline(VddBinStr2Enum(powerPin))(site) / CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion
                            dbl_Efuse_round_number = Round(dbl_quotient)
                            
                            '''//Round up the value because C651 Efuse must round up the IDS values to fuse.
                            If ((CDec(dbl_Efuse_round_number) - CDec(dbl_quotient)) >= 0) Then
                                dbl_Efuse_round_number = dbl_Efuse_round_number
                            Else
                                dbl_Efuse_round_number = dbl_Efuse_round_number + 1
                            End If
                            
                            '''//Generate the offline value for Efuse processed IDS.
                            '''Note: Efuse processed IDS = Efuse number * Efuse resolution.
                            '''For project with Efuse DSP vbt code.
                            ids_for_offline(VddBinStr2Enum(powerPin))(site) = dbl_Efuse_round_number * CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion '''unit: mA
                            
                            '''//According to the current BinCut testJob, update the roundup number of offline IDS values to Efuse category.
                            '''20210716: Modified to decide Efuse setReadDecimal or setWriteDecimal for generate_offline_IDS_IGSim_Parallel.
                            If getBinCutJobDefinition(bincutJobName) > dict_EfuseCategory2BinCutTestJob.Item(str_IDS_PowerDomain) Then '''Efuse setReadDecimal
                                '''For project with Efuse DSP vbt code.
                                CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Read.Decimal(site) = dbl_Efuse_round_number
                            Else '''Efuse setWriteDecimal
                                '''For project with Efuse DSP vbt code.
                                CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Write.Decimal(site) = dbl_Efuse_round_number
                            End If
                        End If
                    Next site
                    
                    '''//Generate offline Efuse product voltage of each p_mode in powerSeq() of BinCut powerDomain.
                    strAry_powerSeq = BinCut_Power_Seq((VddBinStr2Enum(powerPin))).Power_Seq
                    
                    For j = 0 To UBound(strAry_powerSeq)
                        p_mode = VddBinStr2Enum(strAry_powerSeq(j))
                
                        If AllBinCut(p_mode).Used = True Then
                            '''//Get Efuse category for IDS.
                            str_Efuse_read_pmode = get_Efuse_category_by_BinCut_testJob("read", VddBinName(p_mode))
                        
                            '''//Only testJobs of BinCut search can generate Dynamic_IDS_Zone for each mode.
                            '''Note: If str_Efuse_read_pmode is empty, it means that Efuse product voltages not fused before the current testJob.
                            If str_Efuse_read_pmode <> "" And ids_for_offline(VddBinStr2Enum(powerPin))(site) <> 0 Then
                                For Each site In TheExec.sites
                                    Find_IDS_ZONE_per_site ids_for_offline(VddBinStr2Enum(powerPin)), p_mode
                                    Generate_IDS_ZONE_Voltage_Per_Site ids_for_offline(VddBinStr2Enum(powerPin)), p_mode
                                    Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site p_mode
                                    
                                    If siteLng_final_PassBin(site) > 1 Then
                                        idx_step = DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(siteLng_final_PassBin(site), 1)
                                    Else '''bin1
                                        idx_step = (Rnd * DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(siteLng_final_PassBin(site), 1)) / 1
                                        '''ToDo: maybe we can add the vbt code to check voltage_inheritance of each p_mode for the powerDomain.
                                    End If
                                    
                                    dbl_generated_BV = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                                    
                                    '''//Calculate the Efuse product voltage for p_mode.
                                    '''20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
                                    dbl_generated_Product = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step)
                                    
                                    '''//Update Efuse product voltages (GradeVdd) for each p_mode to Efuse data structure.
                                    '''ToDo: Remember to check "EFUSE_BitDef_Table" for UDRP, UDRE, and CFG.
                                    '''For project with Efuse obj vbt code.
                                    '''20210705: Modified to find Efuse Category for the BinCut performance mode by checking the current BinCut testjob.
                                    '''20210714: Modified to generated offline Efuse product voltage.
                                    If str_Efuse_read_pmode <> "" Then
                                        If LCase(AllBinCut(p_mode).powerPin) Like "vdd_pcpu" Then
                                            '''For project with Efuse DSP vbt code.
                                            UDRP_Fuse.Category(UDRP_Index(str_Efuse_read_pmode)).Read.Value(site) = dbl_generated_Product
                                        ElseIf LCase(AllBinCut(p_mode).powerPin) Like "vdd_ecpu" Then
                                            '''For project with Efuse DSP vbt code.
                                            UDRE_Fuse.Category(UDRE_Index(str_Efuse_read_pmode)).Read.Value(site) = dbl_generated_Product
                                        Else
                                            '''For project with Efuse DSP vbt code.
                                            CFGFuse.Category(CFGIndex(str_Efuse_read_pmode)).Read.Value(site) = dbl_generated_Product
                                        End If
                                    End If
                                Next site
                            End If '''If IDS_for_offline(VddBinStr2Enum(powerPin))(site) <> 0
                        End If '''If AllBinCut(p_mode).Used = True
                    Next j '''each p_mode in powerSeq()
                ElseIf dict_IsCorePowerInBinCutFlowSheet.Item(powerPin) = False Then '''OtherRail
                    '''//Use IDS limit as the upper bound of random number.
                    '''//Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
                    '''20210715: Modified to align IDS offline simulation with unit: mA.
                    '''20210721: Modified to use gb_IDS_hi_limit(VddBinStr2Enum(powerDomain), passBinCut) for OtherRail.
                    upperbound = gb_IDS_hi_limit(VddBinStr2Enum(pinGroup_BinCut(i)), siteLng_final_PassBin(site))  '''unit: mA
                    
                    '''//Get lowerbound.
                    '''CorePower has the Efuse category about IDS name, and Efuse always rounds up DC_TEST IDS values as 1# resolution at least.
                    '''Use Efuse resolution as lowerbound for random numbers.
                    If got_Efuse_category = True Then
                        '''For project with Efuse DSP vbt code.
                        lowerbound = 1# * CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion '''unit: mA
                    Else
                        lowerbound = 0
                    End If
                    
                    '''//Check if lowerbound and upperbound are correct.
                    If lowerbound > upperbound Then
                        tempbound = lowerbound
                        lowerbound = upperbound
                        upperbound = tempbound
                    End If
                    
                    '''//Generate the random value for IDS.
                    For Each site In TheExec.sites
                        '''//Ratio for IDS offline simulation.
                        If str_Efuse_read_ProductIdentifier <> "" Then
                            dbl_ratio_Ids = 0.3
                        Else
                            dbl_ratio_Ids = 0.85
                        End If
                        ids_for_offline(VddBinStr2Enum(powerPin))(site) = UserSiteVar(str_IDS_PowerDomain, upperbound, lowerbound) * dbl_ratio_Ids '''unit: mA
                        
                        If got_Efuse_category = True Then
                            '''//Update offline IDS values to Efuse data structure if str_IDS_PowerDomain is the correct Efuse category.
                            '''//For Efuse IDS category, it should round up DC_TEST IDS values with Efuse resolution.
                            '''For project with Efuse DSP vbt code.
                            dbl_quotient = ids_for_offline(VddBinStr2Enum(powerPin))(site) / CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion
                            dbl_Efuse_round_number = Round(dbl_quotient)
                            
                            If ((CDec(dbl_Efuse_round_number) - CDec(dbl_quotient)) >= 0) Then ''''MUST have
                                dbl_Efuse_round_number = dbl_Efuse_round_number
                            Else
                                dbl_Efuse_round_number = dbl_Efuse_round_number + 1
                            End If
                                                        
                            '''//Generate the offline value for Efuse processed IDS.
                            '''Note: Efuse processed IDS = Efuse number * Efuse resolution.
                            '''For project with Efuse DSP vbt code.
                            ids_for_offline(VddBinStr2Enum(powerPin))(site) = dbl_Efuse_round_number * CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion '''unit: mA
                            
                            '''//According to the current BinCut testJob, update the roundup number of offline IDS values to Efuse category.
                            If getBinCutJobDefinition(bincutJobName) > dict_EfuseCategory2BinCutTestJob.Item(str_IDS_PowerDomain) Then '''Efuse setReadDecimal
                                '''For project with Efuse DSP vbt code.
                                CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Read.Decimal(site) = dbl_Efuse_round_number
                            Else '''Efuse setWriteDecimal
                                '''For project with Efuse DSP vbt code.
                                CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Write.Decimal(site) = dbl_Efuse_round_number
                            End If
                        End If
                    Next site
                End If '''If dict_IsCorePowerInBinCutFlowSheet.item(powerPin) = True
            Else
                '''ToDo: for the non-BinCut powerDomain, maybe it can generate random values here...
                str_IDS_PowerDomain = ""
                got_Efuse_category = False
                TheExec.Datalog.WriteComment powerPin & " is not BinCut CorePower or OtherRail for generate_offline_IDS_IGSim_Parallel. Error!!!"
                TheExec.ErrorLogMessage powerPin & " is not BinCut CorePower or OtherRail for generate_offline_IDS_IGSim_Parallel. Error!!!"
            End If
        End If '''If powerPin <> ""
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of generate_offline_IDS_IGSim_Parallel"
    If AbortTest Then Exit Function Else Resume Next
End Function
