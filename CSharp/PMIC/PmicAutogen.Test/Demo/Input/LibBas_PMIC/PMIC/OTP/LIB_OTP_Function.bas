Attribute VB_Name = "LIB_OTP_Function"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

Public g_DictOTPNameIndex As New Dictionary
'Public g_DictOTPAddrIndex As New Dictionary
'Public gDictFileDataIndex As New Dictionary

'Private m_sPreTmpEnWrd As String   ''''20200313
Public g_alOTPDefValue() As Long
Public g_alDefReal() As Long       ''''20200313
Public g_alDefRealUpdate() As Long ''''20200313
Public g_bOtpEnable As Boolean

Private m_avOtpDataInfo() As Variant
Public g_lDebugDumpCnt As Long

'DefaultReal
Public gDW_RealDef_fromMAP As New DSPWave
Public gDW_RealDef_fromWrite As New DSPWave
'Public gDW_RealDef_CompareMAPWrite As New DSPWave

Public g_Total_OTP, g_Total_AHB, g_OTP_With_AHB As Long


Public Function ArrangeOtpTable(Optional r_sSheetName As String = "OTP_register_map")
    Dim sFuncName As String: sFuncName = "ArrangeOtpTable"
    On Error GoTo ErrHandler
    
    '====================================================
    '=  You can adjust the width of each column         =
    '=  from the following constant value assigned      =
    '====================================================
    Const OTP_REGISTER_NAME_WIDTH = 55
    Const NAME_WIDTH = 47
    Const REG_NAME_WIDTH = 85
    Const OWNER_WIDTH = 18
    Const INST_NAME_WIDTH = 20
    Const BIT_WIDTH = 15
    Const COMMENT_WIDTH = 35 ''''=Description
    Const DEFAULT_OR_REAL_WIDTH = 15
    Const END_WIDTH = 12
    ''''----------------------------------------------
    Const BLACKCOLOR = 1
    Const WHITECOLOR = 2
    Const REDCOLOR = 3
    Const LIGHTGREENCOLOR = 4
    Const BLUECOLOR = 5
    Const YELLOWCOLOR = 6
    Const PINKCOLOR = 7
    Const CYANCOLOR = 8
    Const GREENCOLOR = 10
    Const PURPLECOLOR = 13
    Const DARKCYANCOLOR = 14
    Const GREYCOLOR = 15
    Const ORANGECOLOR = 16
    ''''----------------------------------------------
    Const DEVICE_WIDTH = 20
    Const FONTNAME = "Calibri"  ''''"Calibri" or "Arial"
    ''''----------------------------------------------

    Dim lColIdx As Long
    Dim objSheet As Object, ObjOriginCell As Object, ObjCellRow As Object, ObjCellColumn As Object
    Dim lRowCnt As Long, lColCnt As Long
    Dim sRange As String
    Dim lCellWidth As Long
    Dim lFontColor As Long
    Dim lDeviceFlag As Long
    Dim alFailFlag(9999) As Long
    Dim sCellStr As String
    Dim lColComment As Long
    '========================================================

    'r_sSheetName = Sheets(1).Name
    Worksheets(r_sSheetName).Select
    Set objSheet = ActiveSheet
    ActiveWindow.Zoom = 80
      
     '2018/03/26-----------------------
'    Worksheets(r_sSheetName).Unprotect
    Worksheets(r_sSheetName).Select
    Range("A1").Select
    With Selection
         .Font.Name = FONTNAME
         .Font.Size = 12
         .Font.Bold = True
         .ColumnWidth = OTP_REGISTER_NAME_WIDTH
         .Interior.ColorIndex = 0
         .Font.ColorIndex = REDCOLOR
         .HorizontalAlignment = xlCenter 'new
     End With
    '----------------------------------
    
    '======== Set Font and size for all cell ==========
    Range("a1:az4000").Select
    With Selection
        .Font.Name = FONTNAME
        .Font.Size = 11
        .Font.Bold = False
        .RowHeight = 18
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = False ''Can NOT wraptext
    End With
    '===================================================
    
    '====== Initialization ==========
    lRowCnt = 0
    lColCnt = 0
    '===================================
    
    Range("a3").Select
    Set ObjOriginCell = ActiveCell
    Set ObjCellRow = ObjOriginCell
    
    Do Until (LCase(Trim(ObjCellRow.Value)) = "end")
        lRowCnt = lRowCnt + 1
        If LCase(ObjCellRow.Value) Like "otp*register*name*" Then

            sRange = "" & CStr(lRowCnt) & ":" & CStr(lRowCnt)
            Range(sRange).Select
            With Selection
              .WrapText = False ''True, Can NOT wraptext
              .RowHeight = 20
            End With
            '== adjust the width of each column ===
            Set ObjCellColumn = ObjCellRow: lColCnt = 0
            
            Do
                lDeviceFlag = 0
                sCellStr = LCase(Trim(ObjCellColumn.Value))
                If sCellStr Like "otp*register*name*" Then
                    lCellWidth = OTP_REGISTER_NAME_WIDTH
                    'lcolumnBitDef = lColCnt
                ElseIf sCellStr = "reg_name" Then
                    lCellWidth = NAME_WIDTH + OTP_REGISTER_NAME_WIDTH
                ElseIf sCellStr = "name" Then
                    lCellWidth = NAME_WIDTH
                ElseIf (sCellStr Like "otp*owner*") Then
                    lCellWidth = OWNER_WIDTH
                ElseIf sCellStr = "*inst*name*" Then
                    lCellWidth = INST_NAME_WIDTH
                ElseIf (sCellStr Like "*comment*") Then
                    lCellWidth = COMMENT_WIDTH
                    lColComment = lColCnt
                ElseIf sCellStr Like "default*or*real*" Or sCellStr Like "*default*" Or sCellStr Like "*real*" Then
                    lCellWidth = DEFAULT_OR_REAL_WIDTH
               ElseIf sCellStr = "bw" Or sCellStr = "idx" Or sCellStr = "bw" Or sCellStr = "otp_b0" Or sCellStr = "otp_a0" Or sCellStr = "otpreg_add" Or sCellStr = "otpreg_ofs" Then
                    lCellWidth = BIT_WIDTH
                ElseIf sCellStr = "end" Then
                    lCellWidth = END_WIDTH
                Else
                    lCellWidth = DEVICE_WIDTH
                    lDeviceFlag = 1
                End If
                 
                ObjCellColumn.Select
                With Selection
                    .Font.Name = FONTNAME
                    .Font.Size = 12
                    .Font.Bold = True
                    .ColumnWidth = lCellWidth
                    .Interior.ColorIndex = BLUECOLOR
                    .Font.ColorIndex = WHITECOLOR
                    .HorizontalAlignment = xlCenter 'new

                    If (sCellStr Like "*end*") Then
                        .Font.Name = FONTNAME
                        .Font.Size = 11
                        .Font.Bold = True
                        .Interior.ColorIndex = GREENCOLOR 'BlueColor
                    End If
  
                End With
                           
                Set ObjCellColumn = ObjCellColumn.Offset(0, 1)
                lColCnt = lColCnt + 1 'This is used to record totally how many column in this OTP
            Loop Until (IsEmpty(ObjCellColumn.Value))
            
        ElseIf (Not (IsEmpty(ObjCellRow.Value))) Then
            '== some specific keyword with different color ===
            Set ObjCellColumn = ObjCellRow
            
            For lColIdx = 0 To lColCnt - 1
                sCellStr = LCase(Trim(ObjCellColumn.Value))
                If (sCellStr = "default*") Then
                    lFontColor = DARKCYANCOLOR
                ElseIf (sCellStr = "real") Then
                    lFontColor = PINKCOLOR
                ElseIf (sCellStr = "trim") Then
                    lFontColor = PURPLECOLOR
                ElseIf (ObjCellColumn.Value Like "*(F*)*") Then
                    alFailFlag(lColIdx) = 1
                    ObjCellColumn.Select
                    With Selection
                        .Interior.ColorIndex = REDCOLOR
                    End With
                Else
                    lFontColor = BLACKCOLOR
                End If
                
                ObjCellColumn.Select
                With Selection
                    .Font.Name = FONTNAME
                    .Font.Size = 11
                    .Font.ColorIndex = lFontColor
                    .HorizontalAlignment = xlCenter ' xlLeft 'xlCenter
                    If (lColIdx = lColComment) Then
                        .HorizontalAlignment = xlLeft
                    ElseIf lColIdx < 4 Then
                        .HorizontalAlignment = xlLeft
                    End If
                End With

                Set ObjCellColumn = ObjCellColumn.Offset(0, 1)
            Next lColIdx
    
        End If
        
        Set ObjCellRow = ObjCellRow.Offset(1, 0)
        
        '___Freeze panes to always see the head line and test name
        ActiveWindow.FreezePanes = False
        objSheet.Select
        
        '20180920 re-set the clean up range
        '-------------------------------------------------------
        'objSheet.Range("A1:AT60000").ClearOutline
        
        Dim lUsedRow, lUsedCol As Long
        lUsedRow = objSheet.UsedRange.Rows.Count
        lUsedCol = objSheet.UsedRange.Columns.Count
        objSheet.Activate
        ActiveSheet.Range(Cells(4, 1), Cells(lUsedRow, lUsedCol)).ClearOutline
        '-------------------------------------------------------
        objSheet.Range("B4").Select
        ActiveWindow.FreezePanes = True
        'Worksheets(r_sSheetName).Protect
        Exit Function
    Loop     'Do Until (Lcase(ObjCellRow.Value) = "end")
    
Table_end:

    ''''Set Interior Color to Blue on the 'End' Row
    If (LCase(Trim(ObjCellRow.Value)) = "end") Then
        ''Debug.Print "It's End."
        Set ObjCellColumn = ObjCellRow
        ObjCellColumn.Select
        With Selection
            .Font.ColorIndex = WHITECOLOR
            .Font.Name = FONTNAME
            .Font.Size = 11
            .Font.Bold = True
            .Interior.ColorIndex = GREENCOLOR 'BlueColor
            .HorizontalAlignment = xlLeft
            .RowHeight = 20
        End With
    End If
    

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function InitializeOtpTable()
    Dim sFuncName As String: sFuncName = "InitializeOtpTable"
    On Error GoTo ErrHandler
    
        '___Parse_OTP_Table only if g_DictOTPNameIndex.Count = 0
        Call InitializeOtpData
        
        '___Create dictionary based on OTP_register_name (g_DictOTPNameIndex)
        '___Used for SearchOtpIdxByName, looking up the index by OTP_register_name
        Call CreateDictionary_ByOtpName
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ShowOtpInitialLog()
    '___Datalog the OTP_register_map parsing result/OTPAddToDataIndexCreate/OTPRev/OTPRevDataUpdate/AHB_Init
    Dim sFuncName As String: sFuncName = "ShowOtpInitialLog"
    On Error GoTo ErrHandler
    Dim lParseOTPTable As Long
    Dim lDictOTPNameIndex As Long
    If g_DictOTPNameIndex.Count <> 0 Then lParseOTPTable = -1 '-1 means true; Create dictionary after finsish the parsing work
    If g_DictOTPNameIndex.Count <> 0 Then lDictOTPNameIndex = -1
    If InStr(UCase(g_sOTPRevisionType), "V") > 0 Then g_bOTPRevDataUpdate = True

    '___Datalog
    TheExec.Flow.TestLimit lParseOTPTable, -1, -1, TName:="parse_OTP_Table:" & g_sOTP_SHEETNAME
    TheExec.Flow.TestLimit lDictOTPNameIndex, -1, -1, TName:="auto_OTPNameIndexCreate"
    TheExec.Flow.TestLimit g_bOTPRevDataUpdate, -1, -1, TName:="auto_OTPRevDataUpdate:" & g_sOTPRevisionType
    TheExec.Flow.TestLimit g_lOTPRevision, TName:="auto_OTPRev"
    
    ''''20200313, move out to InitializeOtp()
''''    '20200212 TTR from MP7P,20200313
''''    If g_bTTR_ALL = False Then
''''        '___Step1:Initialized OTP Category Result & OTPRegAddrData [MUST]
''''        '___Init OTPdata Read/Write/DefaultToReal/AHB_ReadVal/AHB_ReadVal_ByMaskOfs
''''        Call InitializeOtpDataElement '(EraseFlag:=Not (gB_OTP_Debug))
''''        TheExec.Flow.TestLimit 1, 1, 1, TName:="init_OTPData_Element"
''''    End If
         
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SearchOtpIdxByName(ByVal sKeyName As String) As Integer
    'Input OTP_Register_Name and get back the index
    Dim sFuncName As String: sFuncName = "SearchOtpIdxByName"
    On Error GoTo ErrHandler

    If Not g_DictOTPNameIndex.Exists(sKeyName) Then
        TheExec.ErrorLogMessage "<Error> SearchOtpIdxByName: " & sKeyName & " not found."
    GoTo ErrHandler:
    Else
        SearchOtpIdxByName = CInt(g_DictOTPNameIndex(sKeyName))
    End If

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20200313 update
'___Original function name is auto_OTPCategoryResult_Initialize
Public Function InitializeOtpDataElement() '(Optional EraseFlag As Boolean = True)
    '___Init OTPdata Read/Write/DefaultToReal/AHB_ReadVal/AHB_ReadVal_ByMaskOfs
    Dim sFuncName As String: sFuncName = "InitializeOtpDataElement"
    On Error GoTo ErrHandler
    Dim lOtpIdx As Long
    'Dim mL_sitecnt As Long
    Dim lTempInit As Long
    'Dim lBitWidth As Long
    'Dim lHexLen As Long
    'Dim lCatSize As Long
    
    'mL_sitecnt = TheExec.Sites.Existing.Count - 1
    
    'DefaultReal
    'lCatSize = UBound(g_OTPData.Category) + 1
    
    ''''20200313, move to InitializeOtpData()
''''    For Each Site In TheExec.Sites
''''        gDW_RealDef_fromMAP.CreateConstant 0, lCatSize, DspLong
''''        gDW_RealDef_fromWrite.CreateConstant 0, lCatSize, DspLong
''''    Next Site

''''    Dim alDefReal() As Long
''''    ReDim alDefReal(UBound(g_OTPData.Category))
    
    Dim tResetOTPCat As g_tOTPCategoryParamResultSyntax
    
    ''''By this way to let all Fuse parameters to Nothing/Clear
    ''''Could choose any one of the members as the representative (.Decimal, .HexStr, ...)
    Set tResetOTPCat.Value = Nothing
    
    For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313 AHB New Method
        g_OTPData.Category(lOtpIdx).Read = tResetOTPCat
        
        ''''20200313
        If (g_alDefReal(lOtpIdx) = 1) Then ''''Real case (Trim)
            g_OTPData.Category(lOtpIdx).Write = tResetOTPCat ''''20200313 add, but copy default value must do it again.
        End If
        
        g_OTPData.Category(lOtpIdx).svAhbReadVal = lTempInit
        g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs = lTempInit
        
        ''''20200313 mask
''''        '20190121
''''        lBitWidth = g_OTPData.Category(lOtpIdx).lBitWidth
''''        lHexLen = IIf((lBitWidth Mod (4)) > 0, (lBitWidth \ 4) + 1, lBitWidth \ 4)
''''        g_OTPData.Category(lOtpIdx).Write.HexStr = "0x" + Right("00000000" & Hex(g_OTPData.Category(lOtpIdx).lDefaultValue), lHexLen)

        ''''unused
        'If UCase(g_OTPData.Category(lOTPIdx).sDefaultorReal) = UCase("Default") Then g_OTPData.Category(lOTPIdx).CheckDefaultReal = ""
        'If UCase(g_OTPData.Category(lOTPIdx).sDefaultorReal) = UCase("Real") Then g_OTPData.Category(lOTPIdx).CheckDefaultReal = "NeedToUpdateRealValue"
        
        
        ''''20200313, move to InitializeOtpData()
''''        'DefaultReal
''''        If LCase(g_OTPData.Category(lOtpIdx).sDefaultorReal) Like "*real*" Then '
''''            alDefReal(lOtpIdx) = 1
''''            'gDW_RealDef_fromMAP.Element(lOTPIdx) = 1
''''            'Else 'If g_OTPData.Category(lOTPIdx).sDefaultorReal Like "*Default*" Then
''''            'gDW_RealDef_fromMAP.Element(lOTPIdx) = 0  'default or blank
''''        End If
    Next lOtpIdx
    
    ''''20200313, move to InitializeOtpData()
''''    For Each Site In TheExec.Sites
''''        gDW_RealDef_fromMAP.Data = alDefReal  'gDW_RealDef_fromMAP -> Real=1 , Default=0
''''    Next Site
    
    
    TheExec.Flow.TestLimit 1, 1, 1, TName:=sFuncName
      
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___Includes normal case set write value to (OTP Write Value) and Write AHB Register in FW OTP mode
'___TTR Result: This function takes ~72us @6 site
Public Function auto_OTPCategory_SetWriteDecimal(ByVal sCateName As String, ByVal slSetWriteValue As SiteLong) As Variant

    Dim sFuncName As String: sFuncName = "auto_OTPCategory_SetWriteDecimal"
    On Error GoTo ErrHandler
    Dim vDecimal As Variant
    Dim lBitWidth As Long
    Dim sDLogStr As String
    Dim sDefaultORReal As String
    Dim lOtpIdx As Long
    Dim sTempHeader As String
    Dim sTempEnd As String
    Dim sDLogTrace As String
    Dim sbSaveSiteStatus As New SiteBoolean
    Dim sbCom As New SiteBoolean
    
    sbSaveSiteStatus = TheExec.Sites.Selected
    lOtpIdx = SearchOtpIdxByName(sCateName)
    lBitWidth = g_OTPData.Category(lOtpIdx).lBitWidth
    
   '20200225 TTR The Logical compare contribute more than 470us
    sbCom = slSetWriteValue.BitWiseAnd(Not (2 ^ lBitWidth - 1))
    TheExec.Sites.Selected = sbCom.LogicalNot
    
    'sbCom = slSetWriteValue.Compare(GreaterThanOrEqualTo, 2 ^ lBitWidth).LogicalOr(slSetWriteValue.Compare(LessThan, 0))
    'If (sbCom.Any(True)) Then
    If TheExec.Sites.Selected.Count = 0 Then
        TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out. " + sCateName + " exceed " & 2 ^ lBitWidth & " or  < 0"
        TheExec.Sites.Selected = sbSaveSiteStatus
        Exit Function
        'GoTo ErrHandler
    End If
    
    
    If g_bTTR_ALL = True Then
        '___20200224 Move the OTPData2DSPWave to BurnWriteOTP
    g_OTPData.Category(lOtpIdx).Write.Value = slSetWriteValue

    '___20200224 Update default/real to array but not dspwave
         g_alDefRealUpdate(lOtpIdx) = 1
    
    Else

        '20190819 Moises: If the owner is system or design, should not be overwitten
        If LCase(g_OTPData.Category(lOtpIdx).sOTPOwner) Like "*design*" Then
            TheExec.Datalog.WriteComment "[Warning] The OTP Owner is 'design', should not be overwitten !!!  Please check " & sCateName & """ !!!   Instance :" & TheExec.DataManager.InstanceName
    '        Call TheExec.Flow.TestLimit(ResultVal:=1, TName:="auto_OTPCategory_SetWriteDecimal_OTP_Owner_Check", formatStr:="%0.0f", lowVal:=0, hiVal:=0, TNum:=999)
    '        Exit Function
        ElseIf LCase(g_OTPData.Category(lOtpIdx).sOTPOwner) Like "*system*" Then
            TheExec.Datalog.WriteComment "[Warning] The OTP Owner is 'system', should not be overwitten !!!  Please check " & sCateName & """ !!!   Instance :" & TheExec.DataManager.InstanceName
    '        Call TheExec.Flow.TestLimit(ResultVal:=1, TName:="auto_OTPCategory_SetWriteDecimal_OTP_Owner_Check", formatStr:="%0.0f", lowVal:=0, hiVal:=0, TNum:=999)
    '        Exit Function 'Janet : Nop design/system restriction till new otp map release (asked by Moises)
        End If
        
        sDefaultORReal = UCase(g_OTPData.Category(lOtpIdx).sDefaultORReal)
        
        g_alDefRealUpdate(lOtpIdx) = 1 '20200224
        '---------------------------------------------------------

        ''Notice: DefaultorReal
                sTempEnd = ""
                If sDefaultORReal = UCase("Default") Then
                    sTempEnd = ",Index:" & FormatLog(lOtpIdx, -4) & ", <Default Value> "
                Else
                    sTempEnd = ",Index:" & FormatLog(lOtpIdx, -4) & ",                 "
                End If

        'Only Real will be updated

        If sDefaultORReal = UCase("Default") Then

            TheExec.Datalog.WriteComment "<CHECK_DefaulReal> " & sCateName & " OTP Register is default, but someone try to update default value, skip it!!!!"

        ElseIf sDefaultORReal = UCase("Real") Then

                For Each Site In TheExec.Sites.Selected

                '''Stpe0: Convert the Real Value to Long
                vDecimal = CLng(slSetWriteValue(Site))

                ''''20190618 mask
                'Call auto_OTPData2DSP(CLng(m_vDecimal), site, StartPoint, EndPoint, g_OTPData.Category(lOtpIdx).BitWidth, g_OTPData.Category(lOtpIdx).wBitIndex)

                '___Update the Write Value to OTPData Category and convert decimal value to Hex and Bin 20190520
                g_OTPData.Category(lOtpIdx).Write.Value(Site) = vDecimal
                g_OTPData.Category(lOtpIdx).Write.HexStr(Site) = "0x" + Hex(vDecimal)
                g_OTPData.Category(lOtpIdx).Write.BitStrM(Site) = ConvertFormat_Dec2Bin_Complement(vDecimal, lBitWidth)

                If (g_bSetWriteDebugPrint) Then 'otp_template 20190330
                    sDLogTrace = "b" & g_OTPData.Category(lOtpIdx).Write.BitStrM(Site)
                   
                    sTempHeader = FormatLog("OTPData SetWriteDecimal", -35)
                    sDLogStr = "Site(" + CStr(Site) + ") " + sTempHeader + FormatLog(sCateName, CInt(g_lOTPCateNameMaxLen)) + " = " + FormatLog(vDecimal, -10)
                    sDLogStr = sDLogStr + sTempEnd
                    sDLogStr = sDLogStr + "," + Trim(sDLogTrace)
                    TheExec.Datalog.WriteComment sDLogStr
            End If
            Next Site

            '___Update the write value to the proper position on gD_wPGMData
            'Call RunDsp.LocateOTPData2gDw(slSetWriteValue, lBitWidth, g_OTPData.Category(lOtpIdx).wBitIndex) '20200224 comment out
        End If
    End If
    
TheExec.Sites.Selected = sbSaveSiteStatus

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'___20190627 GetWriteDecimal and Write AHB in addition if FW OTP is enabled or OTP_Burn enable word is not selected
'___20200211 TTR Record (3 Cat. on 1 site: 1*10^-4s =>9*10^-5)~10us improved
Public Function auto_OTPCategory_GetWriteDecimal(r_sCateName As String, _
                                                 ByRef r_slGetWriteVal As SiteLong, _
                                                 Optional r_bUpdateAHB As Boolean = False) As Boolean
    Dim sFuncName As String: sFuncName = "auto_OTPCategory_GetWriteDecimal"
    On Error GoTo ErrHandler
    Dim vDecimal As Variant
    Dim sDLogStr As String
    Dim sDefaultORReal As String
    Dim lOtpIdx As Integer
    Dim sTempHeader As String
    Dim sTempEnd As String

    lOtpIdx = SearchOtpIdxByName(r_sCateName)

    '___Normal case
    r_slGetWriteVal = g_OTPData.Category(lOtpIdx).Write.Value

    If g_bTTR_ALL = False Then
        ''Notice: DefaultorReal
        If (g_bGetWriteDebugPrint) Then
            'lBitWidth = g_OTPData.Category(lOtpIdx).lBitWidth ''''20200313, unused
            sDefaultORReal = g_OTPData.Category(lOtpIdx).sDefaultORReal
            sTempEnd = ""
            If sDefaultORReal = UCase("Default") Then
                sTempEnd = ",Index:" & FormatLog(lOtpIdx, -4) & ", <Default Value> "
            End If
            For Each Site In TheExec.Sites
                vDecimal = CLng(r_slGetWriteVal(Site))
                sTempHeader = FormatLog("OTPData GetWriteDecimal", -35)
                sDLogStr = "Site(" + CStr(Site) + ") " + sTempHeader + FormatLog(r_sCateName, CInt(g_lOTPCateNameMaxLen)) + " = " + FormatLog(vDecimal, -10)
                sDLogStr = sDLogStr + sTempEnd
                TheExec.Datalog.WriteComment sDLogStr
            Next Site
        End If
    End If

    '___Write AHB in addition if r_bUpdateAHB is true
    If r_bUpdateAHB = True Then
        AHB_WRITE CLng(g_OTPData.Category(lOtpIdx).sAhbAddress), r_slGetWriteVal, g_OTPData.Category(lOtpIdx).lCalDeciAhbByMaskOfs
    End If
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''___Debug purpose
'Public Function auto_OTPCategory_GetWriteDecimal_DefaulValue(r_sCateName As String, ByRef mSiteLongValue_Default As SiteLong, _
'                                                                    Optional mDebugPrintLog As Boolean = True) As Boolean
'On Error GoTo ErrHandler
'    Dim sFuncName As String: sFuncName = "auto_OTPCategory_GetWriteDecimal_DefaulValue"
'
'    Dim vDecimal As Variant
'    Dim lBitWidth As Long
'    Dim sDLogStr As String
'    Dim sDefaultORReal As String
'
'    Dim i, k As Integer
'
'    Dim sTempHeader As String
'
'    auto_OTPCategory_GetWriteDecimal_DefaulValue = False
'
'     For Each Site In TheExec.Sites
'             'TheExec.Datalog.WriteComment ""
'            ''Step1:Get the OTPDataIndex
'            i = SearchOtpIdxByName(r_sCateName)
'
'            lBitWidth = g_OTPData.Category(i).lBitWidth
'            sDefaultORReal = g_OTPData.Category(i).sDefaultorReal
'
'            With g_OTPData.Category(i)
'               vDecimal = .lDefaultValue
'               mSiteLongValue_Default(Site) = vDecimal
'
'            ''Notice: DefaultorReal
'               Dim sTempEnd As String
'                sTempEnd = ""
'               If sDefaultORReal = UCase("Default") Then
'                   sTempEnd = ",Index:" & FormatLog(i, -4) & ", <Default Value> "
'               End If
'
'            End With
'
'            '''Stpe2: Update the Value to OTPData Category
'            'g_OTPData.Category(i).RealValue(Site) = vDecimal
'            If (mDebugPrintLog) Then
'                sTempHeader = FormatLog("OTPData GetWriteDecimal DEFAULT", -35)
'                sDLogStr = "Site(" + CStr(Site) + ") " + sTempHeader + FormatLog(r_sCateName, CInt(g_lOTPCateNameMaxLen)) + " = " + FormatLog(vDecimal, -10)
'                sDLogStr = sDLogStr + sTempEnd
'                TheExec.Datalog.WriteComment sDLogStr
'            End If
'
'
'    Next Site
'
'    auto_OTPCategory_GetWriteDecimal_DefaulValue = True
'Exit Function
'
'Error:
'TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out. @@@ " + r_sCateName + ":" + CStr(vDecimal) + " > " + CStr(2 ^ lBitWidth)
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function
''___Debug Purpose
'Public Function auto_OTPCategory_GetWriteDecimal_TrimValue(r_sCateName As String, ByRef mSiteLongValue_Real As SiteLong, _
'                                                                    Optional mDebugPrintLog As Boolean = True) As Boolean
'On Error GoTo ErrHandler
'    Dim sFuncName As String: sFuncName = "auto_OTPCategory_GetWriteDecimal_TrimValue"
'
'    Dim vDecimal As Variant
'    Dim lBitWidth As Long
'    Dim sDLogStr As String
'    Dim sDefaultORReal As String
'
'    Dim Site As Variant
'    Dim i, k As Integer
'
'    Dim sTempHeader As String
'
'    auto_OTPCategory_GetWriteDecimal_TrimValue = False
'
'     For Each Site In TheExec.Sites
'             'TheExec.Datalog.WriteComment ""
'            ''Step1:Get the OTPDataIndex
'            i = SearchOtpIdxByName(r_sCateName)
'
'            lBitWidth = g_OTPData.Category(i).lBitWidth
'            mSiteLongValue_Real = g_OTPData.Category(i).Read.Value(Site)
'
'            With g_OTPData.Category(i)
'               vDecimal = g_OTPData.Category(i).Read.Value(Site)
'
'            ''Notice: DefaultorReal
'               Dim sTempEnd As String
'                sTempEnd = ""
'               If sDefaultORReal = UCase("Default") Then
'                   sTempEnd = ",Index:" & FormatLog(i, -4) & ", <Default Value> "
'               End If
'
'            End With
'
'            '''Stpe2: Update the Value to OTPData Category
'            'g_OTPData.Category(i).RealValue(Site) = vDecimal
'            If (mDebugPrintLog) Then
'                sTempHeader = FormatLog("OTPData GetWriteDecimal REAL", -35)
'                sDLogStr = "Site(" + CStr(Site) + ") " + sTempHeader + FormatLog(r_sCateName, CInt(g_lOTPCateNameMaxLen)) + " = " + FormatLog(vDecimal, -10)
'                sDLogStr = sDLogStr + sTempEnd
'                TheExec.Datalog.WriteComment sDLogStr
'            End If
'
'
'    Next Site
'
'    auto_OTPCategory_GetWriteDecimal_TrimValue = True
'Exit Function
'
'Error:
'TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out. @@@ " + r_sCateName + ":" + CStr(vDecimal) + " > " + CStr(2 ^ lBitWidth)
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'___If AHB Reg was overwritten. Get (OTP Read Value) and update it to corresponding AHB register.
Public Function auto_OTPCategory_GetReadDecimal_AHB(r_sCateName As String, ByRef r_slGetAHBVal As SiteLong, _
                                                                    Optional r_bDebugPrintLog As Boolean = True) As Boolean
    Dim sFuncName As String: sFuncName = "auto_OTPCategory_GetReadDecimal_AHB"
    On Error GoTo ErrHandler
    Dim lOtpIdx As Long
    Dim sAHBRegName As String, lAHBOffset As Long, lAHBBW As Long, lAHBIdx As Long
    Dim slCalcReadAHBOTP As New SiteLong
    Dim sTempHeader As String, mS_tempEnd As String, sDLogStr As String

    
    lAHBIdx = -999
       
    lOtpIdx = SearchOtpIdxByName(r_sCateName)
    'Step1: Get AHB information from OTP r_sCateName
    With g_OTPData.Category(lOtpIdx)
         sAHBRegName = .sRegisterName
         'lAHBBW = .lBitWidth
         lAHBOffset = .lOtpRegOfs
    End With
    
    If g_OTPData.Category(lOtpIdx).sAhbAddress <> "NA" Then
        lAHBIdx = g_dictCRCByAHBRegName.Item(sAHBRegName)
    
        If TheExec.TesterMode = testModeOffline Then
            r_slGetAHBVal = g_OTPData.Category(lOtpIdx).Write.Value
        Else
            With g_OTPData.Category(lOtpIdx)
                'AHB_READ_ByAddr CLng(.sAhbAddress), r_slGetAHBVal, r_bDebugPrintLog
                AHB_READ CLng(.sAhbAddress), r_slGetAHBVal ', .lCalDeciAhbByMaskOfs, r_bDebugPrintLog
            End With
        End If
    
        '___Calc. AHB Data
        g_OTPData.Category(lOtpIdx).svAhbReadVal = r_slGetAHBVal
        '.svAhbReadValByMaskOfs = r_slGetAHBVal.ShiftRight(lAHBOffset).BitwiseAnd(2 ^ lAHBBW - 1)

'        For Each Site In TheExec.Sites
'        With g_OTPData.Category(lOtpIdx)
'                .svAhbReadValByMaskOfs(Site) = r_slGetAHBVal.ShiftRight(lAHBOffset)
'                .svAhbReadValByMaskOfs(Site) = (.svAhbReadValByMaskOfs(Site)) And (.lCalDeciAhbByMaskOfs)
'                r_slGetAHBVal(Site) = .svAhbReadValByMaskOfs(Site)
'        End With
'        Next Site
        ''''20200313, need to check [remove site-loop] 2020/03/28 OK!
        If TheExec.TesterMode = testModeOnline Then
            With g_OTPData.Category(lOtpIdx)
                .svAhbReadValByMaskOfs = r_slGetAHBVal.ShiftRight(lAHBOffset).BitWiseAnd(.lCalDeciAhbByMaskOfs)
                r_slGetAHBVal = .svAhbReadValByMaskOfs
            End With
        End If


        'Step3: Datalog AHB Value
        slCalcReadAHBOTP = g_OTPData.Category(lOtpIdx).svAhbReadValByMaskOfs
    Else
        ''''20200313 add
        TheExec.Datalog.WriteComment sFuncName + ": Can not find AHBIndex!!!!"
        GoTo ErrHandler
    End If


    If (r_bDebugPrintLog) Then
        For Each Site In TheExec.Sites
            sTempHeader = FormatLog("AHBData GetReadDecimal", -35)
            sDLogStr = "Site(" + CStr(Site) + ") " + sTempHeader + FormatLog(r_sCateName, CInt(g_lOTPCateNameMaxLen)) + " = " + FormatLog(slCalcReadAHBOTP, -10)
            sDLogStr = sDLogStr + mS_tempEnd
            TheExec.Datalog.WriteComment sDLogStr
        Next Site
    End If
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___Write (OTP write value) to corresponding AHB register [store Value to OTP and AHB Write]
'___20200313 update, mask some unused variables
Public Function auto_OTPCategory_SetWriteDecimal_AHB(r_sCateName As String, ByVal slWriteAHBVal As SiteLong, _
                                                       Optional r_bDebugPrintLog As Boolean = True) As Boolean
    Dim sFuncName As String: sFuncName = "auto_OTPCategory_SetWriteDecimal_AHB"
    On Error GoTo ErrHandler

    Dim lOtpIdx As Long
    'Dim sCateName As String, sAHBRegName As String, lAHBOffset As Long, lAHBBW As Long, lAHBIdx As Long
    Dim slWrite2AHB As New SiteLong
    Dim sTempHeader As String, mS_tempEnd As String, sDLogStr As String

'    auto_OTPCategory_SetWriteDecimal_AHB = False
'
'    If TheExec.Sites.Selected.Count = 0 Then
'        TheExec.Datalog.WriteComment "TheExec.Sites.Selected.Count = 0"
'        Exit Function
'    End If

    lOtpIdx = SearchOtpIdxByName(r_sCateName)
    'Step1: Get AHB information from OTP r_sCateName
'    With g_OTPData.Category(lOtpIdx)
'       sCateName = .sOtpRegisterName
'       sAHBRegName = .sRegisterName
'       lAHBBW = .lBitWidth
'       lAHBOffset = .lOtpRegOfs
'    End With
          
    'Step2: Write AHB and OTP: '20200313 need to verify it again
    If g_OTPData.Category(lOtpIdx).sAhbAddress <> "NA" Then
        slWrite2AHB = slWriteAHBVal
        With g_OTPData.Category(lOtpIdx)
            .Write.Value = slWriteAHBVal '20190627 Toppy Check here
            'AHB_WRITEDSC_ByAddr CLng(.sAhbAddress), slWrite2AHB
            AHB_WRITE CLng(.sAhbAddress), slWrite2AHB, 255 - g_OTPData.Category(lOtpIdx).lCalDeciAhbByMaskOfs
        End With
    Else
        slWrite2AHB = -999
        ''''20200313 add
        TheExec.Datalog.WriteComment sFuncName + ": Can not find AHB Address!!!!"
        GoTo ErrHandler
    End If
    
    'Step3: Datalog AHB Vale
    If (r_bDebugPrintLog) Then
        mS_tempEnd = ""
        For Each Site In TheExec.Sites
            sTempHeader = FormatLog("AHBData SetWriteDecimal", -35)
            sDLogStr = "Site(" + CStr(Site) + ") " + sTempHeader + FormatLog(r_sCateName, CInt(g_lOTPCateNameMaxLen)) + " = " + FormatLog(slWrite2AHB, -10)
            sDLogStr = sDLogStr + mS_tempEnd
            TheExec.Datalog.WriteComment sDLogStr
        Next Site
    End If

    'auto_OTPCategory_SetWriteDecimal_AHB = True
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'
'Public Function auto_OTPCategory_SetReadDecimal(r_sCateName As String, ByVal mSiteLongValue_ReadBack As SiteLong) As Variant ', _
'                                           Optional r_bDebugPrintLog As Boolean = True) As Variant
'On Error GoTo ErrHandler
'    Dim sFuncName As String: sFuncName = "auto_OTPCategory_SetReadDecimal"
'
'    'Dim vDecimal As Variant
'    Dim mL_BitWidth As Long
'    Dim sDLogStr As String
'    'Dim mS_DefaultORReal As String
'
'    'Dim Site As Variant
'    Dim i As Integer
'    Dim sTempHeader As String
'    '----------------------------------------------------
'    ''Step1:Get the OTPDataIndex
'    i = SearchOtpIdxByName(r_sCateName)
'    mL_BitWidth = g_OTPData.Category(i).lBitWidth
'    'mS_DefaultORReal = g_OTPData.Category(i).sDefaultorReal
'    '----------------------------------------------------
'     For Each Site In TheExec.Sites
'        With g_OTPData.Category(i)
'           Call auto_Read_OTPCatResultFromRegAddrData(i, .lOtpA0, .lOtpB0, .lBitWidth, .lOtpOffset)
'           mSiteLongValue_ReadBack(Site) = .Read.Value(Site)
'           'mSiteLongValue_ReadBack(Site) = vDecimal
'
'
'           ''Notice: for no capture data:
'           Dim mS_tempEnd As String
'           mS_tempEnd = ""
'           If .Read.BitStrM(Site) = "" Then
'               mS_tempEnd = ",Index:" & FormatLog(i, -4) & ", Error: No Capture Data!!!  BitStrM=" & .Read.BitStrM(Site)
'           End If
'
'        End With
'
'        '''Stpe2: Update the Value to OTPData Category
'        'g_OTPData.Category(i).RealValue(Site) = vDecimal
'        If (OTP_PARM.SetRead_DebugPrint) Then
'            sTempHeader = FormatLog("OTPData SetReadDecimal", -35)
'            sDLogStr = "Site(" + CStr(Site) + ") " + sTempHeader + FormatLog(r_sCateName, CInt(g_lOTPCateNameMaxLen)) + " = " + FormatLog(g_OTPData.Category(i).Read.Value(Site), -10)
'            sDLogStr = sDLogStr + mS_tempEnd
'            TheExec.Datalog.WriteComment sDLogStr
'        End If
'    Next Site
'
'Exit Function
'
'Error:
'TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out. @@@ " + r_sCateName + ":" + CStr(g_OTPData.Category(i).Read.Value(Site)) + " > " + CStr(2 ^ mL_BitWidth)
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'___This function is used to get OTP Read and AHBWrite (if r_bUpdateAHB = True)
'___20200313 update, and mask some unused codes
Public Function auto_OTPCategory_GetReadDecimal(r_sCateName As String, _
                                           ByRef r_slGetReadVal As SiteLong, _
                                           Optional r_bUpdateAHB As Boolean = False) As Boolean
    Dim sFuncName As String: sFuncName = "auto_OTPCategory_GetReadDecimal"
    On Error GoTo ErrHandler

    Dim vDecimal As Variant
    Dim lBitWidth As Long
    Dim sDLogStr As String
    Dim sDefaultORReal As String
    Dim lOtpIdx As Integer
    Dim sTempHeader As String
    Dim sTempEnd As String

    'auto_OTPCategory_GetReadDecimal = False
    '----------------------------------------------------
    ''Step1:Get the OTPDataIndex
    lOtpIdx = SearchOtpIdxByName(r_sCateName)
    '----------------------------------------------------
    
    '___GetRead Part
    r_slGetReadVal = g_OTPData.Category(lOtpIdx).Read.Value

        '''Stpe2: Update the Value to OTPData Category
    If (g_bGetReadDebugPrint) Then
        lBitWidth = g_OTPData.Category(lOtpIdx).lBitWidth
        sDefaultORReal = g_OTPData.Category(lOtpIdx).sDefaultORReal

        For Each Site In TheExec.Sites
            vDecimal = r_slGetReadVal(Site)
            sTempEnd = ""
            If g_OTPData.Category(lOtpIdx).Read.BitStrM(Site) = "" Then
                sTempEnd = ",Index:" & FormatLog(lOtpIdx, -4) & ", Error: No Capture Data!!!  BitStrM=" & g_OTPData.Category(lOtpIdx).Read.BitStrM(Site)
            End If
            
            sTempHeader = FormatLog("OTPData GetReadDecimal", -35)
            sDLogStr = "Site(" + CStr(Site) + ") " + sTempHeader + FormatLog(r_sCateName, CInt(g_lOTPCateNameMaxLen)) + " = " + FormatLog(vDecimal, -10)
            sDLogStr = sDLogStr + sTempEnd
            TheExec.Datalog.WriteComment sDLogStr
        Next Site
    End If
    
    '___Write AHB in addition if r_bUpdateAHB is true
    If r_bUpdateAHB = True Then
        'AHB_WRITEDSC_ByAddr CLng(g_OTPData.Category(lOTPIdx).sAhbAddress), r_slGetReadVal, g_OTPData.Category(lOTPIdx).lCalDeciAhbByMaskOfs
        AHB_WRITE CLng(g_OTPData.Category(lOtpIdx).sAhbAddress), r_slGetReadVal, g_OTPData.Category(lOtpIdx).lCalDeciAhbByMaskOfs
    End If

    'auto_OTPCategory_GetReadDecimal = True
Exit Function

''''20200313, mask
''''Error:
''''TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out. @@@ " + r_sCateName + ":" + CStr(vDecimal) + " > " + CStr(2 ^ lBitWidth)
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'
'Public Function OTP_SimTestCase()
'
'On Error GoTo ErrHandler
'    Dim sFuncName As String: sFuncName = "OTP_SimTestCase"
'
'    Dim slSimValue As New SiteLong
'    Dim slGetSimValue As New SiteLong
'    TheExec.Datalog.WriteComment ""
'    TheExec.Datalog.WriteComment "<" + sFuncName + ">" + ":" + sFuncName
'
'
'     For Each Site In TheExec.Sites
'       slSimValue(Site) = 4 + Site
'     Next Site
'
'    ''Store the Value:
'    Call auto_OTPCategory_SetWriteDecimal(r_sCateName:="OTP_LDO_ADC_TRIM_1531", mSiteLongValue:=mSL_SimValue) ', SetWrite_DebugPrint:=True)
'
'    ''Readback the Value: Must set "Default or Real'=Read in the table.
'    Call auto_OTPCategory_GetWriteDecimal(r_sCateName:="OTP_LDO_ADC_TRIM_1531", r_slGetReadVal:=mSL_GetSimValue) ', mDebugPrintLog:=True)
'
'    'TheExec.Datalog.WriteComment ""
'
'Exit Function
'
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'''' 20170926 Add check IEDAString
'''' use to prevent the hang up issue if the iEDA register key has the wrong format.
Public Function CheckNCombineIEDA(r_sIEDA As String) As String
    Dim sFuncName As String: sFuncName = "CheckNCombineIEDA"
    On Error GoTo ErrHandler
    Dim iSiteIdx As Integer
    Dim asIeda() As String
    Dim sTmpStr As String
    Dim iIedaSize As Integer

    asIeda = Split(r_sIEDA, ",")
    iIedaSize = UBound(asIeda) + 1
    
    ''Debug.Print "asIeda size =" & iIedaSize
    sTmpStr = ""
    If (iIedaSize < TheExec.Sites.Existing.Count) Then
        If (iIedaSize = 0) Then
            sTmpStr = "NA"
            For iSiteIdx = 1 To (TheExec.Sites.Existing.Count - iIedaSize - 1)
                sTmpStr = sTmpStr + ",NA"
            Next iSiteIdx
        Else
            For iSiteIdx = 1 To (TheExec.Sites.Existing.Count - iIedaSize)
                sTmpStr = sTmpStr + ",NA"
            Next iSiteIdx
        End If
        TheExec.Datalog.WriteComment sFuncName + ": original = " + r_sIEDA
        r_sIEDA = r_sIEDA + sTmpStr
        TheExec.Datalog.WriteComment Space(23) + "  update = " + r_sIEDA + " ......could have the site sequence problem (case1) !!!" + vbCrLf
    
    ElseIf (iIedaSize > TheExec.Sites.Existing.Count) Then ''''should not have this case
        sTmpStr = ""
        For iSiteIdx = 0 To TheExec.Sites.Existing.Count - 1
            If (iSiteIdx = (TheExec.Sites.Existing.Count - 1)) Then
                sTmpStr = sTmpStr + asIeda(iSiteIdx)
            Else
                sTmpStr = sTmpStr + asIeda(iSiteIdx) + ","
            End If
        Next iSiteIdx
        TheExec.Datalog.WriteComment sFuncName + ": original = " + r_sIEDA
        r_sIEDA = sTmpStr
        TheExec.Datalog.WriteComment Space(23) + "  update = " + r_sIEDA + " ......could have the site sequence problem (case2) !!!" + vbCrLf

    End If
    
    CheckNCombineIEDA = r_sIEDA
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CreateDictionary_ByOtpName()
    '___Create dictionary according to OTP_register_name (g_DictOTPNameIndex)
    Dim sFuncName As String: sFuncName = "CreateDictionary_ByOtpName"
    On Error GoTo ErrHandler

    Dim lOtpIdx As Long
    Dim lIndex As Long
    Dim sOTPRegName As String
    Dim sKeyName As String
    Dim vObj As Variant
    
    g_DictOTPNameIndex.RemoveAll
    For lOtpIdx = 0 To (g_Total_OTP - 1) ' was UBound(g_OTPData.Category) '___20200313, AHN New Method
        lIndex = g_OTPData.Category(lOtpIdx).lOtpIdx
        If lOtpIdx <> lIndex Then GoTo ErrHandler
        
        sOTPRegName = g_OTPData.Category(lOtpIdx).sOtpRegisterName
        '___Create g_DictOTPNameIndex
        sKeyName = UCase(sOTPRegName)
        vObj = lIndex
        If g_DictOTPNameIndex.Exists(sKeyName) Then
            g_DictOTPNameIndex.Remove (sKeyName)
        End If
        g_DictOTPNameIndex.Add sKeyName, vObj
        
        g_lOTPCateNameMaxLen = IIf(g_lOTPCateNameMaxLen > Len(sOTPRegName), g_lOTPCateNameMaxLen, Len(sOTPRegName))
    Next lOtpIdx

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function InitializeOtpData()
    Dim sFuncName As String: sFuncName = "InitializeOtpData"
    On Error GoTo ErrHandler
    Dim wsOTPTable As Worksheet
    Dim wsAHBTable As Worksheet
    Dim lOtpBitStrStart As Long
    Dim lOtpIdx As Long
    Dim lStartRow As Long
    Dim lEndCol As Long
    Dim lCatSize As Long
    Dim lOTPMapCol As Long
    Dim lCateRevCnt As Long
    Dim lV1Col As Long

    Dim alIdx() As Long
    Dim wIdx As New DSPWave
    Dim lBitIdx As Long

    'for LocateColnRow_Ahb()
    Dim lRegNameCol As Long
    Dim lRegAddrCol As Long
    Dim lRegIdxCol As Long
    Dim lFieldWidthCol As Long
    Dim lFieldOffsetCol As Long
    Dim lLastRow As Long
    
    If (IsSheetExists(g_sOTP_SHEETNAME) = False) Or (IsSheetExists(g_sAHB_SHEETNAME) = False) Then
        GoTo ErrHandler
    Else
        '___Define AHB_register_map sheet name
        Set wsAHBTable = Sheets(g_sAHB_SHEETNAME)
        '____Define OTP_register_map sheet name
        Set wsOTPTable = Sheets(g_sOTP_SHEETNAME)
    End If
    
    ''''20200313, add here for new AHB method
    Call LocateColnRow_Ahb(lRegNameCol, lRegAddrCol, lRegIdxCol, lFieldWidthCol, lFieldOffsetCol, lLastRow)
    
    ''wsOTPTable.Activate '20200313, it's activate inside the function
    Call LocateColnRow_Otp(lStartRow, lEndCol, lCatSize, lV1Col)
    
    ''''20200313, New AHB method
    Call Check_OTP_AHB_Category_Size(lCatSize, lLastRow) '2020/03/09 check category size for ALL OTP & AHB
    
    '___OTPData
    'ReDim g_OTPData.Category(m_lCatSize - 1)
    ReDim g_OTPData.Category(g_Total_OTP + g_Total_AHB - g_OTP_With_AHB - 1)
    
    ReDim g_alDefReal(lCatSize - 1)
    ReDim g_alDefRealUpdate(lCatSize - 1)
    '___OTPRev
    ReDim g_OTPRev.Category(lEndCol - lV1Col - 1)
    
    'lCateRevCnt = 0
    ReDim g_OTPRev.Category(lCateRevCnt).DefaultorReal(lCatSize - 1)
    ReDim g_OTPRev.Category(lCateRevCnt).DefaultValue(lCatSize - 1)

    
    If g_bTTR_ALL = True Then
        For Each Site In TheExec.Sites.Existing
             gDW_RealDef_fromMAP.CreateConstant 0, lCatSize, DspLong
             gDW_RealDef_fromWrite.CreateConstant 0, lCatSize, DspLong
        Next Site
    End If
    
    For lOTPMapCol = 1 To lEndCol - 1
        For lOtpIdx = 0 To lCatSize - 1
            Select Case lOTPMapCol
                Case 1
                    g_OTPData.Category(lOtpIdx).sOtpRegisterName = CStr(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 2
                    g_OTPData.Category(lOtpIdx).sName = CStr(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 3
                    g_OTPData.Category(lOtpIdx).sInstanceName = CStr(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 4
                    g_OTPData.Category(lOtpIdx).sRegisterName = CStr(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 5
                    g_OTPData.Category(lOtpIdx).sOTPOwner = CStr(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 6
                    g_OTPData.Category(lOtpIdx).lDefaultValue = CLng(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 7
                    g_OTPData.Category(lOtpIdx).lBitWidth = CLng(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 8
                    g_OTPData.Category(lOtpIdx).lOtpIdx = CLng(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 9
                    g_OTPData.Category(lOtpIdx).lOtpOffset = CLng(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 10
                    g_OTPData.Category(lOtpIdx).lOtpB0 = CLng(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 11
                     g_OTPData.Category(lOtpIdx).lOtpA0 = CLng(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 12
                    g_OTPData.Category(lOtpIdx).lOtpRegAdd = CStr(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 13
                    g_OTPData.Category(lOtpIdx).lOtpRegOfs = CStr(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                Case 14
                    g_OTPData.Category(lOtpIdx).sDefaultORReal = CStr(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol).Value)
                    'DefaultReal, 20200313 add
                    If LCase(g_OTPData.Category(lOtpIdx).sDefaultORReal) Like "*real*" Then '
                        g_alDefReal(lOtpIdx) = 1
                    End If
                
                Case 15
                    '___OTP-AHB Compare
                    '___Put AHB info into OTPData structure right away once gets the register_name
                    Call InitializeOtpDataFromAhbMap(g_OTPData.Category(lOtpIdx).sRegisterName, lOtpIdx)
                   
                '___define BitStr_Start/End
                Case 16
                    '***lOTPBitStrStart = OTP_ADD_A0*OTP_BitsPerADDR +  OTP_BLOCK_B0*OTP_BitsPerBlock + OTP_OFFSET
                    g_OTPData.Category(lOtpIdx).lOtpBitStrStart = ((g_OTPData.Category(lOtpIdx).lOtpA0 * g_iOTP_DATA_BW) _
                                                           + (g_OTPData.Category(lOtpIdx).lOtpB0 * g_iOTP_BITS_PERBLOCK) + g_OTPData.Category(lOtpIdx).lOtpOffset)
                Case 17
                    '***otp_BitStr_End = lOTPBitStrStart + BitWidth - 1
                    g_OTPData.Category(lOtpIdx).lOtpBitStrEnd = g_OTPData.Category(lOtpIdx).lOtpBitStrStart + g_OTPData.Category(lOtpIdx).lBitWidth - 1
                    If (True) Then
                        ''''20200313 update
                        With g_OTPData.Category(lOtpIdx)
                            For Each Site In TheExec.Sites.Existing
                                wIdx.CreateRamp .lOtpBitStrStart, 1, .lBitWidth, DspLong
                                .wBitIndex = wIdx.Copy
                            Next Site
                        End With
                    Else
                        ReDim alIdx(g_OTPData.Category(lOtpIdx).lBitWidth - 1)
                        lOtpBitStrStart = g_OTPData.Category(lOtpIdx).lOtpBitStrStart
                        For lBitIdx = 0 To g_OTPData.Category(lOtpIdx).lOtpBitStrEnd - g_OTPData.Category(lOtpIdx).lOtpBitStrStart
                            alIdx(lBitIdx) = lOtpBitStrStart
                            lOtpBitStrStart = lOtpBitStrStart + 1
                        Next lBitIdx
                        
                        For Each Site In TheExec.Sites.Existing '20190618
                            wIdx.CreateConstant 0, g_OTPData.Category(lOtpIdx).lOtpBitStrEnd - g_OTPData.Category(lOtpIdx).lOtpBitStrStart + 1, DspLong
                            g_OTPData.Category(lOtpIdx).wBitIndex.Data = alIdx
                        Next Site
                    End If

                Case Else
                    For lCateRevCnt = 0 To lEndCol - lV1Col - 1
                         ReDim Preserve g_OTPRev.Category(lCateRevCnt).DefaultorReal(lCatSize - 1)
                         ReDim Preserve g_OTPRev.Category(lCateRevCnt).DefaultValue(lCatSize - 1)
                         If lOtpIdx = 0 Then
                             g_OTPRev.Category(lCateRevCnt).PKGName = CStr(wsOTPTable.Cells(lStartRow, lOTPMapCol + lCateRevCnt).Value)
                             g_OTPRev.Category(lCateRevCnt).Index = lCateRevCnt
                         End If
                         g_OTPRev.Category(lCateRevCnt).DefaultValue(lOtpIdx) = CVar(wsOTPTable.Cells(lOtpIdx + lStartRow + 1, lOTPMapCol + lCateRevCnt).Value)
                         g_OTPRev.Category(lCateRevCnt).DefaultorReal(lOtpIdx) = g_OTPData.Category(lOtpIdx).sDefaultORReal
                    Next lCateRevCnt
                    If lOtpIdx = lCatSize - 1 Then GoTo EndParsing
            End Select
        Next lOtpIdx
    Next lOTPMapCol
    If g_bTTR_ALL = True Then
        TheExec.Flow.TestLimit 1, 1, 1, TName:="init_OTPData_Element"
    End If
          
EndParsing:
    '___20200313 add
    ''''In the wafer sort, the 1st Run may not enable all existing Sites,
    ''''So it's recommended that put it in the all sites loop
    ''''BTW, this call only do once.
    For Each Site In TheExec.Sites.Existing
        gDW_RealDef_fromWrite.CreateConstant 0, lCatSize, DspLong
        gDW_RealDef_fromMAP.CreateConstant 0, lCatSize, DspLong
        gDW_RealDef_fromMAP.Data = g_alDefReal
    Next Site

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Private Function Check_OTP_AHB_Category_Size(OTP_Size As Long, AHB_Last_Row As Long)
    Dim sFuncName As String: sFuncName = "Check_OTP_AHB_Category_Size"
    On Error GoTo ErrHandler

    Dim ws_OTP_col, ws_AHB_col As Integer
    Dim m_wsOTPTable As Worksheet
    Dim m_wsAHBTable As Worksheet
    Dim RegName As String
    Dim BFName As String
    
    g_OTP_With_AHB = 0
    g_dictAHBRegQty.RemoveAll
    
    Set m_wsAHBTable = Sheets(g_sAHB_SHEETNAME)
    Set m_wsOTPTable = Sheets(g_sOTP_SHEETNAME)
    m_wsAHBTable.Activate
    For ws_OTP_col = 2 To OTP_Size
        RegName = UCase(m_wsOTPTable.Cells(ws_OTP_col + 2, 4))
        BFName = UCase(m_wsOTPTable.Cells(ws_OTP_col + 2, 2))
            If RegName <> "" And BFName <> "" Then
                g_dictAHBRegQty(RegName & "_" & BFName) = RegName & "_" & BFName
            End If
        RegName = ""
        BFName = ""
    Next
    
    For ws_AHB_col = 2 To AHB_Last_Row
        RegName = UCase(m_wsAHBTable.Cells(ws_AHB_col + 2, 5))
        BFName = UCase(m_wsAHBTable.Cells(ws_AHB_col + 2, 6))
            If RegName <> "" And BFName <> "" And g_dictAHBRegQty.Exists(RegName & "_" & BFName) = True Then
                g_OTP_With_AHB = g_OTP_With_AHB + 1
            End If
        RegName = ""
        BFName = ""
    Next

    g_Total_OTP = OTP_Size
    g_Total_AHB = AHB_Last_Row - 1
    
    g_dictAHBRegQty.RemoveAll

    '---OTP----AHB---
    '---(O)----(X)--- => OTP_Qty = g_Total_OTP - g_OTP_With_AHB
    '---(X)----(O)--- => AHB_Qty = g_Total_AHB - g_OTP_With_AHB
    '---(O)----(O)--- => OTP_Qty = AHB_Qty = g_OTP_With_AHB
    'OTPData.Category size should be "OTP_Qty + AHB_Qty - g_OTP_With_AHB -1"

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ConvertOtpVersion_String2Value(Optional HOST_INTERFACE As g_eHOST_INTERFACE_TYPE = ePLATFORM_ID, Optional r_sInput As String) As Long
    Dim sFuncName As String: sFuncName = "ConvertOtpVersion_String2Value"
    On Error GoTo ErrHandler
    Dim sPerChar As String

    Dim IIdx As Integer
    Dim lDecodeVal As Long
    Dim asArray() As Variant
    Dim alValArray() As Variant
    Dim bFound As Boolean: bFound = False
   
    
    
       Select Case HOST_INTERFACE
       
           Case ePLATFORM_ID:
                asArray = Array("A33", "B", "C", "D", "E", "F", "G", "H", "I", _
                                  "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                  "S", "T", "U", "V", "W", "X", "Y", "Z", "D33")
            
                alValArray = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, _
                                   9, 10, 11, 12, 13, 14, 15, 16, 17, _
                                  18, 19, 20, 21, 22, 23, 24, 25, 0)
           Case eOTP_CONSUMER_TYPE:
                asArray = Array("Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "HTOL/Qual", "Reserved", _
                                  "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "System Eng.", "Reserved", "Reserved", _
                                  "SoC-SiVal", "Reserved", "Reserved", "PMU-SiVal", "Reserved", "Reserved", "Reserved", "Reserved")
            
                alValArray = Array(0, 0, 0, 0, 0, 0, 0, 3, 0, _
                                   0, 0, 0, 0, 0, 0, 1, 0, 0, _
                                  2, 0, 0, 0, 0, 0, 0, 0)
                                                            
           Case eOTP_REVISION_TYPE:
                asArray = Array("Blank", "Trim only", "C", "D", "E", "F", "G", "H", "I", _
                                  "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                  "S", "T", "U", "V", "W", "X", "Y", "Z")
            
                alValArray = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, _
                                   9, 10, 11, 12, 13, 14, 15, 16, 17, _
                                  18, 19, 20, 21, 22, 23, 24, 25)
                                  
           Case eTP_OTP_VERSION:
                asArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", _
                                  "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                  "S", "T", "U", "V", "W", "X", "Y", "Z")
            
                alValArray = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, _
                                   10, 11, 12, 13, 14, 15, 16, 17, 18, _
                                   19, 20, 21, 22, 23, 24, 25, 26)
            Case Else
             
               TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": Can not find the case.."
            
      End Select
    
    
    r_sInput = UCase(r_sInput)
    lDecodeVal = 0
    sPerChar = UCase(r_sInput)

    For IIdx = 0 To UBound(asArray)
        'If InStr(UCase(asArray(iIdx)), sPerChar) Then
         If (sPerChar = UCase(asArray(IIdx))) Then
           lDecodeVal = alValArray(IIdx)
            bFound = True
           Exit For
        End If
    Next IIdx
    
    
    If bFound = True Then
       ConvertOtpVersion_String2Value = lDecodeVal
    Else
       ConvertOtpVersion_String2Value = 0
       TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": Can not find the match case value."
    End If

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ConvertOtpVersion_Value2String(Optional HOST_INTERFACE As g_eHOST_INTERFACE_TYPE = ePLATFORM_ID, _
                                        Optional r_lInput As Long = 0, _
                                            Optional r_sRtnString As String = "") As String
    '2018/05/23 Claire
    Dim sFuncName As String: sFuncName = "ConvertOtpVersion_Value2String"
    On Error GoTo ErrHandler
    Dim IIdx As Long
    Dim sDecodeString        As String
    Dim asArray()           As Variant
    Dim alValArray()  As Variant
    Dim asLetter() As Variant
    Dim bFound As Boolean: bFound = False
           
    asLetter = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", _
                     "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                     "S", "T", "U", "V", "W", "X", "Y", "Z")
                                  
                                  
     Select Case HOST_INTERFACE
     
         Case ePLATFORM_ID:
              asArray = Array("A33", "B", "C", "D", "E", "F", "G", "H", "I", _
                                "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                "S", "T", "U", "V", "W", "X", "Y", "Z")
                                       
              alValArray = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, _
                         9, 10, 11, 12, 13, 14, 15, 16, 17, _
                        18, 19, 20, 21, 22, 23, 24, 25)
                                
         Case eOTP_CONSUMER_TYPE:
              asArray = Array("Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "HTOL/Qual", "Reserved", _
                                "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "System Eng.", "Reserved", "Reserved", _
                                "SoC-SiVal", "Reserved", "Reserved", "PMU-SiVal", "Reserved", "Reserved", "Reserved", "Reserved")
                         
              alValArray = Array(0, 0, 0, 0, 0, 0, 0, 3, 0, _
                                 0, 0, 0, 0, 0, 0, 1, 0, 0, _
                                2, 0, 0, 0, 0, 0, 0, 0)
                                

         Case eOTP_REVISION_TYPE:
              asArray = Array("Blank", "Trim only", "C", "D", "E", "F", "G", "H", "I", _
                                "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                "S", "T", "U", "V", "W", "X", "Y", "Z")
              alValArray = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, _
                         9, 10, 11, 12, 13, 14, 15, 16, 17, _
                        18, 19, 20, 21, 22, 23, 24, 25)
                                                  

         Case eTP_OTP_VERSION:
              asArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", _
                                "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                "S", "T", "U", "V", "W", "X", "Y", "Z")
                                
              alValArray = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, _
                                 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, _
                                 19, 20, 21, 22, 23, 24, 25, 26)
         Case Else
             TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": Can not find the case.."
          
    End Select
    
    
    For IIdx = 0 To UBound(asArray)
        If r_lInput = CLng(alValArray(IIdx)) Then
           r_sRtnString = asArray(IIdx)
           sDecodeString = asLetter(IIdx)
            bFound = True
           Exit For
        End If
    Next IIdx
    
    If bFound = True Then
       ConvertOtpVersion_Value2String = sDecodeString
    Else
       ConvertOtpVersion_Value2String = ""
       TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": Can not find the match case value."
    End If
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Public Function auto_OTP_String2String(Optional HOST_INTERFACE As g_eHOST_INTERFACE_TYPE = ePLATFORM_ID, Optional r_sInput As String)
Public Function ConvertOtpVersion_String2String(Optional HOST_INTERFACE As g_eHOST_INTERFACE_TYPE = ePLATFORM_ID, Optional r_sInput As String)
    Dim sFuncName As String: sFuncName = "ConvertOtpVersion_String2String"
    On Error GoTo ErrHandler
    
    Dim sPerChar As String
    Dim IIdx As Integer
    Dim sDecodeString As String
    Dim asArray() As Variant
    Dim asLetter() As Variant
    Dim bFound As Boolean: bFound = False
   
    
    asLetter = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", _
                                  "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                  "S", "T", "U", "V", "W", "X", "Y", "Z")
                                  
     Select Case HOST_INTERFACE
     
         Case ePLATFORM_ID:
              asArray = Array("A33", "B", "C", "D", "E", "F", "G", "H", "I", _
                                "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                "S", "T", "U", "V", "W", "X", "Y", "Z", "D33")
                                
             asLetter = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", _
                                "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                "S", "T", "U", "V", "W", "X", "Y", "Z", "A")
          
         Case eOTP_CONSUMER_TYPE:
              asArray = Array("Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "HTOL/Qual", "Reserved", _
                                "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "Reserved", "System Eng.", "Reserved", "Reserved", _
                                "SoC-SiVal", "Reserved", "Reserved", "PMU-SiVal", "Reserved", "Reserved", "Reserved", "Reserved")

                                                          
         Case eOTP_REVISION_TYPE:
              asArray = Array("Blank", "Trim only", "C", "D", "E", "F", "G", "H", "I", _
                                "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                "S", "T", "U", "V", "W", "X", "Y", "Z")

         Case eTP_OTP_VERSION:
              asArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", _
                                "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                                "S", "T", "U", "V", "W", "X", "Y", "Z")
          Case Else
           
             TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": Can not find the case.."
          
    End Select
    
    
    r_sInput = UCase(r_sInput)
    sDecodeString = ""
    sPerChar = UCase(r_sInput)

    For IIdx = 0 To UBound(asArray)
        'If InStr(UCase(asArray(iIdx)), sPerChar) Then
        If (sPerChar = UCase(asLetter(IIdx))) Then
           sDecodeString = asArray(IIdx)
            bFound = True
           Exit For
        End If
    Next IIdx
    
    
    If bFound = True Then
       ConvertOtpVersion_String2String = sDecodeString
    Else
       ConvertOtpVersion_String2String = ""
       TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": Can not find the match case value."
    End If

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckNSetWriteTPVersion()
    Dim sFuncName As String: sFuncName = "CheckNSetWriteTPVersion"
    On Error GoTo ErrHandler
    
    Dim sWorkBookName As String
    Dim slTPOTPVersion As New SiteLong
    
    Dim asWrkBookArray() As String
    
    Dim lTPVersion    As Long
    Dim lTPVersion_M  As Long
    Dim lTPVersion_S  As Long
    
    Dim sTPVersion    As String
    Dim sTPVersion_M  As String
    Dim sTPVersion_S  As String
    
    
    Dim slTPVersion_M  As New SiteLong
    Dim slTPVersion_S  As New SiteLong
    Dim slSVNVersion  As New SiteLong
    Dim slSVNVersion_MSB  As New SiteLong
    Dim slSVNVersion_LSB  As New SiteLong

    '___To get workbook name on IGXL9.0
    If TheExec.SoftwareVersion Like "9.00.00_uflx*" Then
        sWorkBookName = UCase(TheExec.TestProgram.Name)
        sWorkBookName = Replace(sWorkBookName, ".IGXL", "")
    Else
        sWorkBookName = UCase(ActiveWorkbook.Name)
        sWorkBookName = Replace(sWorkBookName, ".XLSM", "")
    End If
        asWrkBookArray = Split(sWorkBookName, "_")
    
    
    If UBound(asWrkBookArray) >= 5 Then
        If Len(asWrkBookArray(5)) = 4 And UCase(Mid(asWrkBookArray(5), 1, 3)) Like "V##" Then   '20171225: add check rule => should be character V concat 2 single digits
            lTPVersion_M = CLng(Mid(asWrkBookArray(5), 2, 2))                                   'asWrkBookArray(5)="V01A" ==>mL_TPVERSION_M="01"  :01-15
            lTPVersion_S = ConvertOtpVersion_String2Value(eTP_OTP_VERSION, Mid(asWrkBookArray(5), 4, 1))  'asWrkBookArray(5)="V01A" ==>mL_TPVERSION_S="A"   :A-O
            If lTPVersion_M > 15 Or lTPVersion_S > 15 Then
                'lTPVersion_M:01-15 (4bits) / lTPVersion_S=A-O (4bits) '2018/07/05
                TheExec.Datalog.WriteComment "WARNING:PLEASE CHECK YOUR TPVERSION_M(" & Mid(asWrkBookArray(5), 2, 2) & ") = " & lTPVersion_M & " < 16."
                TheExec.Datalog.WriteComment "WARNING:PLEASE CHECK YOUR TPVERSION_S(" & Mid(asWrkBookArray(5), 4, 1) & " ) = " & lTPVersion_S & " < 16"
                lTPVersion = 0: sTPVersion = "": sTPVersion_M = "": sTPVersion_S = ""
                 Call MsgBox("WARNING:PLEASE CHECK YOUR TPVERSION_M(" & Mid(asWrkBookArray(5), 2, 2) & ") = " & lTPVersion_M & " < 16." _
                            , vbOKOnly, "WARNING : Please check your test program format.")
                 Call MsgBox("WARNING:PLEASE CHECK YOUR TPVERSION_S(" & Mid(asWrkBookArray(5), 4, 1) & " ) = " & lTPVersion_S & " < 16" _
                            , vbOKOnly, "WARNING : Please check your test program format.")
                 Stop
           Else
            lTPVersion = ConvertFormat_Bin2Dec(ConvertFormat_Dec2Bin_Complement(lTPVersion_M, 4) + ConvertFormat_Dec2Bin_Complement(lTPVersion_S, 4))
            sTPVersion = asWrkBookArray(5): sTPVersion_M = Mid(asWrkBookArray(5), 2, 2): sTPVersion_S = Mid(asWrkBookArray(5), 4, 1)
            End If
        Else
            lTPVersion_M = 0
            lTPVersion_S = 0
            lTPVersion = 0
            sTPVersion = "": sTPVersion_M = "": sTPVersion_S = ""
            TheExec.Datalog.WriteComment ("WARNING:Please check your test program format.")
            Call MsgBox("-@Naming rule:" & vbNewLine & "Device _ Job _ Site# _ OTP version _ Date code _  T/P version_....._.xlsm" _
                        & vbNewLine & "-@EX: " & vbNewLine & "MP3T_CP1_X2_E00_171226_V02D.xlsm" _
                        , vbOKOnly, "WARNING : Please check your test program format.")
            Stop
        End If
    Else
            lTPVersion_M = 0
            lTPVersion_S = 0
            lTPVersion = 0
            TheExec.Datalog.WriteComment ("WARNING:Please check your test program format.")
            Call MsgBox("-@Naming rule:" & vbNewLine & "Device _ Job _ Site# _ OTP version _ Date code _  T/P version_....._.xlsm" _
                        & vbNewLine & "-@EX: " & vbNewLine & "MP3T_CP1_X2_E00_171226_V02D.xlsm" _
                        , vbOKOnly, "WARNING : Please check your test program format.")
            'Stop
    End If
    
    
    For Each Site In TheExec.Sites
        slTPOTPVersion(Site) = lTPVersion
        slTPVersion_M(Site) = lTPVersion_M  'added 2018/03/27
        slTPVersion_S(Site) = lTPVersion_S  'added 2018/03/27
    Next Site
    
    TheExec.Datalog.WriteComment ("TestProgram VERSION: " & sWorkBookName)
    TheExec.Datalog.WriteComment "TP-VERSION =" & sTPVersion & Space(2) & "; TP-VERSION-M =" & sTPVersion_M & Space(2) & "; TP-VERSION-S =" & sTPVersion_S
    
    TheExec.Flow.TestLimit lTPVersion_M, 1, 15, TName:="TPVERSION_M", formatStr:="%.0f"
    TheExec.Flow.TestLimit lTPVersion_S, 1, 15, TName:="TPVERSION_S", formatStr:="%.0f"
    
    
    Call auto_OTPCategory_SetWriteDecimal(g_sOTP_TPVERSION_S, slTPVersion_S)  ', SetWrite_DebugPrint:=True)
    Call auto_OTPCategory_SetWriteDecimal(g_sOTP_TPVERSION_M, slTPVersion_M) ', SetWrite_DebugPrint:=True)
                                                      
    'KC Write SVN revision into OTP per Jeff Cobb Request on 06/18/2018
    
    If UBound(asWrkBookArray) >= 6 Then
        If Len(asWrkBookArray(6)) = 7 And UCase(Mid(asWrkBookArray(6), 1, 7)) Like "SVN####" Then     '20171225: add check rule => should be character SVN concat 4 single digits
            slSVNVersion = CLng(Mid(asWrkBookArray(6), 4, 4))
            
            slSVNVersion_MSB = slSVNVersion
            slSVNVersion_LSB = slSVNVersion.BitWiseAnd(&HFF)
            slSVNVersion_MSB = slSVNVersion.ShiftRight(8).BitWiseAnd(&HFF)
        Else
            slSVNVersion = 0
            slSVNVersion_LSB = 0
            slSVNVersion_MSB = 0
            Call MsgBox("Test Program Naming Rule:" & vbNewLine & "Device _ Job _ Site# _ OTP version _ Date code _  T/P version_SVN Version_.xlsm" _
                        & vbNewLine & "-@EX: " & vbNewLine & "MP4T_CP1_X2_E00_171226_V02D_SNV1234.xlsm" _
                        , vbOKOnly, "WARNING : Please check your test program format.")
        End If
    Else
            Call MsgBox("Test Program Naming Rule:" & vbNewLine & "Device _ Job _ Site# _ OTP version _ Date code _  T/P version_SVN Version_.xlsm" _
                        & vbNewLine & "-@EX: " & vbNewLine & "MP4T_CP1_X2_E00_171226_V02D_SVN1234.xlsm" _
                        , vbOKOnly, "WARNING : Please check your test program format.")
        Stop
    End If
    
    TheExec.Datalog.WriteComment "Test Program SVN Version := " & Format(CStr(slSVNVersion), "0000")
    
    Call auto_OTPCategory_SetWriteDecimal(g_sSVN_VERSION_MSB, slSVNVersion_MSB) ', SetWrite_DebugPrint:=True)
    Call auto_OTPCategory_SetWriteDecimal(g_sSVN_VERSION_LSB, slSVNVersion_LSB) ', SetWrite_DebugPrint:=True)  ''20190129 remove _SPA no LSB information

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function CheckNGetWriteOTPVersion() '(Optional mDebugPrintLog As Boolean = True)
    Dim sFuncName As String: sFuncName = "CheckNGetWriteOTPVersion"
    On Error GoTo ErrHandler
    
    Dim sOTPType As String
    Dim sOTPRevisionType As String
    Dim asTypeArray(2) As String
    Dim alTypeArray(2) As Long
    Dim aslTypeArray(2) As New SiteLong

    'g_sOTPRevisionType
    'g_sOTPRevisionType = "OTP_V" + CStr(g_lOTPRevision) + g_sOTPType    ;EX:OTP_V0/OTP_V1/OTP_V2_APC/OTP_V2_ASC
    'g_sOTPType: "_" + OTP_V0/OTP_V1/OTP_V2_APC/OTP_V2_ASC                     ;APC :PLATFORM_ID/OTP_CONSUMER_TYPE/OTP_REVISION
    sOTPType = g_sOTPType
    sOTPRevisionType = g_sOTPRevisionType

    If sOTPType <> "" Then
        asTypeArray(0) = ConvertOtpVersion_String2String(ePLATFORM_ID, Mid(sOTPType, 1, 1))
        asTypeArray(1) = ConvertOtpVersion_String2String(eOTP_CONSUMER_TYPE, Mid(sOTPType, 2, 1))
        asTypeArray(2) = ConvertOtpVersion_String2String(eOTP_REVISION_TYPE, Mid(sOTPType, 3, 1))
        
        alTypeArray(0) = ConvertOtpVersion_String2Value(ePLATFORM_ID, asTypeArray(0))
        alTypeArray(1) = ConvertOtpVersion_String2Value(eOTP_CONSUMER_TYPE, asTypeArray(1))
        alTypeArray(2) = ConvertOtpVersion_String2Value(eOTP_REVISION_TYPE, asTypeArray(2))
        
        If (g_bOTPRevCheckDebugPrint) Then
            TheExec.Datalog.WriteComment ("")
            TheExec.Datalog.WriteComment ("OTP_RevisionType : " & g_sOTPRevisionType)
            TheExec.Datalog.WriteComment ("OTP_Type         : " & sOTPType)
            TheExec.Datalog.WriteComment ("PLATFORM_ID      : " & Mid(sOTPType, 1, 1) & " >> " & FormatLog(asTypeArray(0), -12) & " =d'" & CStr(alTypeArray(0)))
            TheExec.Datalog.WriteComment ("OTP_CONSUMER_TYPE: " & Mid(sOTPType, 2, 1) & " >> " & FormatLog(asTypeArray(1), -12) & " =d'" & CStr(alTypeArray(1)))
            TheExec.Datalog.WriteComment ("OTP_REVISION_TYPE: " & Mid(sOTPType, 3, 1) & " >> " & FormatLog(asTypeArray(2), -12) & " =d'" & CStr(alTypeArray(2)))
        End If
        
        Call auto_OTPCategory_GetWriteDecimal("OTP_HOST_INTERFACE_PLATFORM_ID_2", aslTypeArray(0))  ', mDebugPrintLog:=True)
        Call auto_OTPCategory_GetWriteDecimal("OTP_HOST_INTERFACE_OTP_CONSUMER_1", aslTypeArray(1)) ', mDebugPrintLog:=True)
        Call auto_OTPCategory_GetWriteDecimal("OTP_HOST_INTERFACE_OTP_REVISION_1", aslTypeArray(2)) ', mDebugPrintLog:=True)
        
        TheExec.Flow.TestLimit aslTypeArray(0), TName:="PLATFORM_ID", formatStr:="%.0f", hiVal:=alTypeArray(0), lowVal:=alTypeArray(0)
        TheExec.Flow.TestLimit aslTypeArray(1), TName:="OTP_CONSUMER_TYPE", formatStr:="%.0f", hiVal:=alTypeArray(1), lowVal:=alTypeArray(1)
        TheExec.Flow.TestLimit aslTypeArray(2), TName:="OTP_REVISION_TYPE", formatStr:="%.0f", hiVal:=alTypeArray(2), lowVal:=alTypeArray(2)
    Else
        If (g_bOTPRevCheckDebugPrint) Then
            TheExec.Datalog.WriteComment ("")
            TheExec.Datalog.WriteComment ("OTP_RevisionType : " & sOTPRevisionType)
            TheExec.Datalog.WriteComment ("OTP_Type         : " & sOTPType)
        End If
    End If


Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'___Used in Auto_AHB_OTP_Write/Read CheckByCondition
Public Function CreateDictionary_ForCompAhbOtp()
    Dim sFuncName As String: sFuncName = "CreateDictionary_ForCompAhbOtp"
    On Error GoTo ErrHandler
    
    Dim lOtpIdx As Long
    Dim lIdx             As Long
    Dim sDefaultORReal     As String, sOTPOwner         As String
    Dim bCheckFlag          As Boolean, lCheckCnt         As Long

    g_DictOTPPreCheckIndex.RemoveAll
    lCheckCnt = 0
    For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
        With g_OTPData.Category(lOtpIdx)
        lIdx = .lOtpIdx
        '--------Check Condition--------------
        sDefaultORReal = UCase(.sDefaultORReal)
        sOTPOwner = UCase(.sOTPOwner)
        '--------Check Condition--------------
        End With
      
        'Check Condition:By g_sOTP_DEFAULT_REAL_FOR_PRECHECK/g_sOTP_DEFAULT_REAL_FOR_PRECHECK
        bCheckFlag = (InStr(UCase(g_sOTP_DEFAULT_REAL_FOR_PRECHECK), sDefaultORReal) > 0) And (InStr(UCase(g_sOTP_OWNER_FOR_CHECK), sOTPOwner) > 0)
        If bCheckFlag Then
            'lCheckCnt = lCheckCnt + 1: mKeyName = lCheckCnt: mValue = lIdx
            lCheckCnt = lCheckCnt + 1
            
            If g_DictOTPPreCheckIndex.Exists(lIdx) Then
                Debug.Print "Index" & lIdx & " Key Exists."
            Else
                g_DictOTPPreCheckIndex.Add lCheckCnt, lIdx
            End If
        End If
    Next lOtpIdx
    'Debug.Print "Total lCheckCnt               =" & lCheckCnt
    'Debug.Print "Total g_DictOTPPreCheckIndex CNT =" & g_DictOTPPreCheckIndex.Count

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___Sub-Function for AHB-OTP comparison,
'___called by "Auto_AHB_OTP_ReadCheck_New" and "Auto_AHB_OTP_WriteCheck_New"
Public Function ReadAhbToCat()
    Dim sFuncName As String: sFuncName = "ReadAhbToCat"
    On Error GoTo ErrHandler
    Dim lOtpIdx As Long

    If TheExec.Sites.Selected.Count = 0 Then
       TheExec.Datalog.WriteComment "TheExec.Sites.Selected.Count = 0"
       Exit Function
    End If
    
    If TheExec.TesterMode = testModeOffline Then Exit Function
    For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
        If g_OTPData.Category(lOtpIdx).sAhbAddress <> "NA" Then
            'AHB_READDSC_ByAddr CLng(g_OTPData.Category(lOTPIdx).sAhbAddress), g_OTPData.Category(lOTPIdx).svAhbReadVal   'Need to debug 20190408
            AHB_READ CLng(g_OTPData.Category(lOtpIdx).sAhbAddress), g_OTPData.Category(lOtpIdx).svAhbReadVal ', OTPData.Category(lOTPIdx).lCalDeciAhbByMaskOfs
        End If
    Next lOtpIdx


Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''___Called by auto_OTP_SetWriteDecimal
''___Set the write value to proper location in gD_wPGMData
'Public Function auto_OTPData2DSP(Val As Long, Site As Variant, ByVal BitStr_start As Long, ByVal BitStr_End As Long, BitWidth As Long, indexWave As DSPWave)
'On Error GoTo ErrHandler
'Dim sFuncName As String: sFuncName = "auto_OTPData2DSP"
'
'Dim BinDspwave As New DSPWave
'Dim ValDSPWave As New DSPWave
'
'    ValDSPWave.CreateConstant Val, 1, DspLong
'    BinDspwave = ValDSPWave.ConvertStreamTo(tldspSerial, BitWidth, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
'    gD_wPGMData.ReplaceElements indexWave, BinDspwave
'
'
'Exit Function
'ErrHandler:
'    RunTimeError sFuncName
'    If AbortTest Then Exit Function Else Resume Next
'End Function

Public Function InitializeOtpGlobalFlag()
    Dim sFuncName As String: sFuncName = "InitializeOtpGlobalFlag"
    On Error GoTo ErrHandler
    If UCase(TheExec.CurrentJob) Like "FT*" Or TheExec.EnableWord("OTP_FTProg") = True Then g_bEnWrdOTPFTProg = True
    
    If g_bEnWrdOTPFTProg = True Then
    'FT need to reset WaferSetup or datalog xy will be the last one's.
        For Each Site In TheExec.Sites
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(Site, -32768)
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(Site, -32768)
        Next Site
    End If
    
    '___OTP_Enable and OTP-FW control
    g_bOTPcmpAHB = TheExec.EnableWord("OTP_cmpAHB")
    g_bOtpEnable = TheExec.EnableWord("A_Enable_OTP_Burn")
    g_bOTPFW = False
    g_bOTPOneShot = False
    g_bFWDlogCheck = False
    g_bOTPOneShot = TheExec.EnableWord("A_Enable_OTP_OneShotBurn") ' oneshot read without otp enable
'    If TheExec.EnableWord("OTP_Expct_Actual_Dlog_EN") = True Then gB_Expct_Actual_Dlog_EN = True
    If g_bOtpEnable Then
        g_bOTPFW = TheExec.EnableWord("A_Enable_OTP_FWBurn")
        g_bFWDlogCheck = TheExec.EnableWord("OTP_FW_Dlog_CHK")
        If (TheExec.EnableWord("A_Enable_OTP_OneShotBurn") = True) And (TheExec.EnableWord("A_Enable_OTP_FWBurn") = True) Then GoTo ErrHandler
        If (TheExec.EnableWord("A_Enable_OTP_FWBurn") = False) And (TheExec.EnableWord("OTP_FW_Dlog_CHK") = True) Then GoTo ErrHandler
    End If
    
    
    '==========Customize the Debug Print Flags Here==========
    '___Print the binary and decimal values of OTP intended to source with DSSC or captured from DSSC
    g_bOTPDsscBitsDebugPrint = False
    
    '___Set True to print all infomation on Datalog, or set False to Datalog the failed items only
    g_bAHBWriteCheckDebugPrint = True
    g_bExpectedActualDebugPrint = True
    g_bSetWriteDebugPrint = True
    g_bGetWriteDebugPrint = False
    g_bGetReadDebugPrint = False
    
    '___Print OTP platformID, consumer type and revision type
    g_bOTPRevCheckDebugPrint = True
    
    g_bDump2CsvDebugPrint = True '20190611
    '==========Customize the Debug Print Flags Here==========
    
    '___Print OTP platformID, consumer type and revision type
    
    '___Log Test Time
    g_bTestTimeProfileDebugPrint = True
    
    
    '---------------------------------------------------------
    '20200313 .Setup.LotSetup.TestMode
    '---------------------------------------------------------
    '0 = A = AEL (Automatic Edge Lock) mode
    '1 = C = Checker mode
    '2 = D = Development/Debug test mode
    '3 = E = Engineering mode (same as Development mode)
    '4 = M = Maintenance mode
    '5 = P = Production test mode
    '6 = Q = Quality Control
    '---------------------------------------------------------
    If TheExec.Datalog.Setup.LotSetup.TESTMODE = 5 Then
        g_bOTPDsscBitsDebugPrint = False
        g_bAHBWriteCheckDebugPrint = False
        g_bSetWriteDebugPrint = False
        g_bGetWriteDebugPrint = False
        g_bGetReadDebugPrint = False
        g_bOTPRevCheckDebugPrint = False
        g_bDump2CsvDebugPrint = False
        g_bTestTimeProfileDebugPrint = False
        g_bExpectedActualDebugPrint = False
        'TheExec.EnableWord("OTP_Production") = True
    End If
        
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function InitializeOtpDataFromAhbMap(r_sRegName As String, r_lOTPIdx As Long) As String
    Dim sFuncName As String: sFuncName = "InitializeOtpDataFromAhbMap"
    On Error GoTo ErrHandler
    '___For OTP-AHB Compare
    Dim wsAHBTable As Worksheet
    Set wsAHBTable = Sheets(g_sAHB_SHEETNAME)
    
    Dim sCellContent As String
    Dim sAHBAddr As String
    Dim lAHBFieldWidth As Long
    Dim lAHBFieldOffset As Long
    Dim iRepeatCnt As Integer
    Dim iBitsIdx As Integer
    
    Dim dRegIdx As Double
    
    Dim lRegNameCol As Long
    Dim lRegAddrCol As Long
    Dim lRegIdxCol As Long
    Dim lFieldWidthCol As Long
    Dim lFieldOffsetCol As Long
    Dim lLastRow As Long
    
    Dim wRtnSerDataVal As New DSPWave
    Dim wRtnParDataVal As Long
    Dim slCalDeciAHBByMaskOfs As New SiteLong

    wsAHBTable.Activate
    Call LocateColnRow_Ahb(lRegNameCol, lRegAddrCol, lRegIdxCol, lFieldWidthCol, lFieldOffsetCol, lLastRow)
    If Not IsError(Application.Match(r_sRegName, wsAHBTable.Range(Cells(2, lRegNameCol), Cells(lLastRow, lRegNameCol)), 0)) Then   'reg name
        '___AHB_Addr
        sCellContent = Application.Match(r_sRegName, wsAHBTable.Range(Cells(2, lRegNameCol), Cells(lLastRow, lRegNameCol)), 0)
        sAHBAddr = Application.Index(wsAHBTable.Range(Cells(2, lRegAddrCol), Cells(lLastRow, lRegAddrCol)), sCellContent)    'reg address
        dRegIdx = Application.Index(wsAHBTable.Range(Cells(2, lRegIdxCol), Cells(lLastRow, lRegIdxCol)), sCellContent)  'index
        g_OTPData.Category(r_lOTPIdx).sAhbAddress = Replace(LCase(sAHBAddr), "0x", "&H")
        '__check how many times the reg_name appears in the column
        iRepeatCnt = Application.CountIf(wsAHBTable.Range(Cells(2, lRegNameCol), Cells(lLastRow, lRegNameCol)), r_sRegName) 'reg name

        '___AHB_Field_Width
        lAHBFieldWidth = g_OTPData.Category(r_lOTPIdx).lBitWidth
        '___AHB_Field_Offset
        lAHBFieldOffset = g_OTPData.Category(r_lOTPIdx).lOtpRegOfs
                   
        '___Could not set it as rundsp function because here is a loop
        Call GetAhbMask(CLng(lAHBFieldWidth), CLng(lAHBFieldOffset), wRtnSerDataVal, wRtnParDataVal) ''''was AHBMask_Calc()
        With g_OTPData.Category(r_lOTPIdx)
        
            .lAhbMaskVal = wRtnParDataVal
            For iBitsIdx = 0 To g_iAHB_BW - 1
            .sAhbMask(iBitsIdx) = (wRtnSerDataVal.Data(iBitsIdx)) 'Xor (1)
            Next iBitsIdx
            slCalDeciAHBByMaskOfs = (wRtnParDataVal) Xor (2 ^ g_iAHB_BW - 1)
            .lCalDeciAhbByMaskOfs = slCalDeciAHBByMaskOfs.ShiftRight(lAHBFieldOffset)
            
        End With
    
    Else
        g_OTPData.Category(r_lOTPIdx).sAhbAddress = "NA"
        '___Init AHB_Mask array with all "1"
        For iBitsIdx = 0 To g_iAHB_BW - 1
            g_OTPData.Category(r_lOTPIdx).sAhbMask(iBitsIdx) = 1
        Next iBitsIdx
    End If

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function init_AHBEnumIdx_to_OTPIndexDict() As Boolean
On Error GoTo ErrHandler
Dim sFuncName As String: sFuncName = "init_AHBEnumIdx_to_OTPIndexDict"

    Dim i As Integer
    Dim RegName As String
    Dim RegBFName As String
    
    'Dim ReferenceTime As Double
    'ReferenceTime = TheExec.Timer
    
    g_dictAHBEnumIdx.RemoveAll
    
    '''20200313, Need to check ''''(g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
    For i = 0 To UBound(g_OTPData.Category)
        If UCase(g_OTPData.Category(i).sRegisterName) <> "" Then
            RegName = UCase(g_OTPData.Category(i).sRegisterName)
            RegBFName = UCase(g_OTPData.Category(i).sRegisterName) & "." & UCase(g_OTPData.Category(i).sName) 'Here to set "REG name" + "." + "BF name"
            If g_dictAHBEnumIdx.Exists(RegName) = False Then
                g_dictAHBEnumIdx(RegName) = i
            End If
            g_dictAHBEnumIdx(RegBFName) = i
        End If
    Next i
    
'Debug.Print TheExec.Timer(ReferenceTime)
Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub LocateColnRow_Ahb(ByRef r_lRegNameCol As Long, ByRef r_lRegAddrCol As Long, ByRef r_lRegIdxCol As Long, _
                       ByRef r_lFieldWidthCol As Long, ByRef r_lFieldOffsetCol As Long, ByRef r_lLastRow As Long)
    Dim sFuncName As String: sFuncName = "LocateColnRow_Ahb"
    On Error GoTo ErrHandler
    '___For OTP-AHB Compare
    '___Parse AHB register information
    Dim wsAHBTable As Worksheet
    Set wsAHBTable = Sheets(g_sAHB_SHEETNAME)

    wsAHBTable.Activate
    
    Dim lRowIdx As Long

    For lRowIdx = 1 To 50
        If LCase(Cells(1, lRowIdx)) Like "reg name" Then
           r_lRegNameCol = lRowIdx
        ElseIf LCase(Cells(1, lRowIdx)) Like "reg address" Then
           r_lRegAddrCol = lRowIdx
        ElseIf LCase(Cells(1, lRowIdx)) Like "index" Then
           r_lRegIdxCol = lRowIdx
        ElseIf LCase(Cells(1, lRowIdx)) Like "field width" Then
           r_lFieldWidthCol = lRowIdx
        ElseIf LCase(Cells(1, lRowIdx)) Like "field offset" Then
           r_lFieldOffsetCol = lRowIdx
        Exit For
        End If
    Next lRowIdx
    '___ find the last row
    Cells(1, r_lRegIdxCol).Select
    r_lLastRow = Selection.End(xlDown).Row

Exit Sub
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub LocateColnRow_Otp(ByRef r_lStartRow As Long, ByRef r_lEndRow As Long, r_lCatSize As Long, ByRef r_lV1Col As Long)
    Dim sFuncName As String: sFuncName = "LocateColnRow_Otp"
    On Error GoTo ErrHandler
    '___Parse OTP register information
    Dim wsOTPTable As Worksheet
    Set wsOTPTable = Sheets(g_sOTP_SHEETNAME)
    Dim lEndRow As Long
    Dim iRowIdx As Integer

    
    wsOTPTable.Activate
    
    '___Look for the start cell
    Range("A1").Select
    'Sometimes it not start from A1
    If Range("A1") = "" Then Selection.End(xlDown).Select
    r_lStartRow = Selection.Row
    
    Cells(r_lStartRow, 1).Select
    Selection.End(xlDown).Select
    lEndRow = Selection.Row
    Cells(r_lStartRow, 1).Select
    Selection.End(xlToRight).Select
    r_lEndRow = Selection.Column
    r_lCatSize = lEndRow - r_lStartRow - 1
    
    For iRowIdx = r_lEndRow To 1 Step -1
        'Usaually, the 1st otp version will be on the right of "Different" column
        If UCase(Cells(r_lStartRow, iRowIdx)) Like "*DIFFERENT*" Then
           r_lV1Col = iRowIdx + 1
           Exit For
        End If
    Next iRowIdx
Exit Sub
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'___20200313, Optimized
Public Function InitializeOtpVersion()
    'Control the OTP version
    Dim sFuncName As String: sFuncName = "InitializeOtpVersion"
    On Error GoTo ErrHandler

    Dim asEnableWords() As String
    Dim asVersion() As String
    Dim lEnWrdsCnt As Long
    Dim iEnWrdsIdx As Integer
    Dim sEnWrd As String
    Dim iEnVersionCnt As Integer
    Dim lIdx As Long
    Dim bmatchFlag As Boolean
    
    'Get Enable Words and judge OTP_Ver
    'lEnWrdsCnt = tl_ExecGetEnableWords(asEnableWords)
    '(otp_template) Merge MPxP/T and MP5T special case 20190318
'    For iEnWrdsIdx = 0 To UBound(asEnableWords)
'        sEnWrd = UCase(CStr(asEnableWords(iEnWrdsIdx)))
'        bmatchFlag = False
        
        sEnWrd = TheExec.CurrentPart
        ''''20200313, only process the enable words with the prefix "OTP_" and it's True
        If (sEnWrd Like "OTP_*") Then
            'If (sEnWrd Like UCase("OTP_V#") Or sEnWrd Like UCase("OTP_V##")) Then 'case "OTP_V1", "OTP_V01"
            If (sEnWrd Like UCase("OTP_ECID_ONLY")) Then 'case "OTP_ECID_ONLY" '20200330 Myint asked !
                bmatchFlag = True
            ElseIf sEnWrd Like UCase("OTP_V#_[A-Z][A-Z][A-Z]") Then 'And TheExec.Flow.EnableWord(sEnWrd) = True) Then 'case MP4T "OTP_V1_AVC"
                bmatchFlag = True
            ElseIf sEnWrd Like UCase("OTP_V##_[A-Z][A-Z][A-Z]") Then 'And TheExec.Flow.EnableWord(sEnWrd) = True) Then 'case "OTP_V01_AVC"
                bmatchFlag = True
            ElseIf sEnWrd Like UCase("OTP_[A-Z][A-Z][A-Z]_V##") Then 'And TheExec.Flow.EnableWord(sEnWrd) = True) Then 'case MP5T/MP9P "OTP_AVC_V01"
                bmatchFlag = True
            ElseIf sEnWrd Like UCase("OTP_[A-Z][A-Z][A-Z]_V#") Then 'And TheExec.Flow.EnableWord(sEnWrd) = True) Then 'case MP5T/MP9P "OTP_AVC_V1"
                bmatchFlag = True
            End If
            
            If (bmatchFlag = True) Then
                iEnVersionCnt = iEnVersionCnt + 1
                lIdx = iEnWrdsIdx
            End If
        End If
'    Next iEnWrdsIdx

    If (iEnVersionCnt = 1) Then
'        sEnWrd = UCase(CStr(asEnableWords(lIdx)))
        
        ''''20200313, If(sEnWrd = g_sOTPVerPrevious), do NOT need to process the below again
        If (sEnWrd = g_sOTPVerPrevious) Then
            g_sOTPVerPrevious = sEnWrd ''''20200313 enable, put here is to keep the static variable
            Call UpdateDefaultValue2gDw(sEnWrd, False) ''''False: only copy the 1st run global array to DSPWave, save TT.
            TheExec.Flow.TestLimit iEnVersionCnt, 1, 1, TName:="OTP_Version_Select_Count", formatStr:="%.0f"
            Exit Function
        End If

        asVersion = Split(sEnWrd, "_")
        
        ''''20200313 update
        ''''--------------------------------------------------
        ''''TP Version case1, V1=1,V2=2, V3=3...etc.
        ''''TP Version case2, V01=1,V02=2, V11=11...etc.
        ''''--------------------------------------------------
        ''''OTP_Revision, A=0,B=1,C=2,D=3,E=4...etc.
        ''''g_sOTPType="_AP(D)" @ "OTP_V2_APD"
        ''''--------------------------------------------------
        g_lOTPRevision = 0
        g_sOTPType = ""
        
        If (sEnWrd Like UCase("OTP_V*_[A-Z][A-Z][A-Z]")) Then
            ''''so asVersion(1) will be like "V*"
            g_lTestProgVersion = CLng(Replace(asVersion(1), "V", ""))
            g_sOTPType = asVersion(2)
            g_lOTPRevision = Asc(Mid(g_sOTPType, 3, 1)) - 65
        ElseIf (sEnWrd Like UCase("OTP_[A-Z][A-Z][A-Z]_V*")) Then
            g_lTestProgVersion = Replace(asVersion(2), "V", "")
            g_sOTPType = asVersion(1)
            g_lOTPRevision = Asc(Mid(g_sOTPType, 3, 1)) - 65
        'ElseIf (sEnWrd Like UCase("OTP_V#") Or sEnWrd Like UCase("OTP_V##")) Then ''''case OTP_V1,OTP_V01
        ElseIf (sEnWrd Like UCase("OTP_ECID_ONLY")) Then  ''''case OTP_ECID_ONLY
            g_lTestProgVersion = 1
            g_sOTPType = "" ''''ECID only
            g_lOTPRevision = 0
        Else
            ''''wrong case
            iEnVersionCnt = -1
            GoTo ErrHandler
        End If

        '20190429 update to Default DSPwave based on the selected OTP version
        g_sOTPRevisionType = "OTP_" + g_sOTPType + "_V" + Format(CStr(g_lTestProgVersion), "00")
        g_sOTPRevisionType = Replace(g_sOTPRevisionType, "__", "_")
        'g_bOTPRevDataUpdate = True
        Call UpdateDefaultValue2gDw(sEnWrd)
        g_sOTPVerPrevious = sEnWrd ''''20200313 enable
        
    ElseIf (iEnVersionCnt = 0) Then
        g_lOTPRevision = 0 '[default]
        g_sOTPVerPrevious = ""
        TheExec.AddOutput "PLEASE SELECT OTP TYPE!!"
        
        ''''20200313, set Fail flag and bin out by BinTable
        ''''if set flag in the flow, here it can be masked
        For Each Site In TheExec.Sites
            TheExec.Sites.Item(Site).FlagState("F_OTP_Init") = logicTrue
        Next Site
    Else
        g_lOTPRevision = 0
        g_sOTPVerPrevious = ""
        TheExec.AddOutput "PLEASE SELECT 'ONE' OTP TYPE ONLY!!"
        
        ''''20200313, set Fail flag and bin out by BinTable
        ''''if set flag in the flow, here it can be masked
        For Each Site In TheExec.Sites
            TheExec.Sites.Item(Site).FlagState("F_OTP_Init") = logicTrue
        Next Site
    End If
    
    TheExec.Flow.TestLimit iEnVersionCnt, 1, 1, TName:="OTP_Version_Select_Count", formatStr:="%.0f"
 
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Dump all default values according to OTPrev to DSPwave -gDW_DefaultDSPRawData
'20200313 update method
Public Function UpdateDefaultValue2gDw(r_sOTPver As String, Optional bUpdateRevDefault As Boolean = True) '20190416
    Dim sFuncName As String: sFuncName = "UpdateDefaultValue2gDw"
    On Error GoTo ErrHandler
    Dim lRevIdx, lOtpIdx As Long
    Dim alBinary() As Long
    Dim wTemp As New DSPWave
    Dim wOTPDefValue As New DSPWave
    Dim wIndex As New DSPWave
    Dim sOTPRevName As String
    Dim wTemp2 As New DSPWave
    Dim m_tmpEW As String
    
    ''''20200313 update, only change OTP Rev (EW) then update the default
    If (bUpdateRevDefault = True) Then
        g_lRevIdx = 999 '20200330 JY in case there's no version to be selected.
        'g_sOTPRevisionType
        For lRevIdx = 0 To UBound(g_OTPRev.Category)
            sOTPRevName = UCase(g_OTPRev.Category(lRevIdx).PKGName)
            'Example:PKGName="OTP_V1"  ;sOTPRevName="V1"
             If r_sOTPver = "OTP_ECID_ONLY" And sOTPRevName = "V1" Then
                g_lRevIdx = 0
                Exit For
             ElseIf UCase(r_sOTPver) = UCase("OTP_" & sOTPRevName) Then '2018/03/26
                g_lRevIdx = g_OTPRev.Category(lRevIdx).Index
                Exit For
             End If
        Next lRevIdx
            
        ''''20190618 update
        wOTPDefValue.CreateConstant 0, CLng(g_iOTP_ADDR_TOTAL) * g_iOTP_REGDATA_BW, DspLong
        For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
            g_OTPData.Category(lOtpIdx).lDefaultValue = g_OTPRev.Category(g_lRevIdx).DefaultValue(lOtpIdx)
            'defVal = g_OTPRev.Category(g_lRevIdx).lDefaultValue(lOTPIdx)
            '20190521 Set init write value as default value
            For Each Site In TheExec.Sites.Existing
               g_OTPData.Category(lOtpIdx).Write.Value = g_OTPRev.Category(g_lRevIdx).DefaultValue(lOtpIdx)
            Next Site
            Call ConvertFormat_Dec2Bin(g_OTPRev.Category(g_lRevIdx).DefaultValue(lOtpIdx), g_OTPData.Category(lOtpIdx).lBitWidth, alBinary)
              
            wIndex.CreateRamp g_OTPData.Category(lOtpIdx).lOtpBitStrStart, 1, g_OTPData.Category(lOtpIdx).lBitWidth, DspLong
            For Each Site In TheExec.Sites '.Existing '20190611
                wTemp.Data = alBinary
                Call wOTPDefValue.ReplaceElements(wIndex, wTemp)
            Next Site
        Next lOtpIdx
        ''''20190618, only do once to get the global value g_alOTPDefValue()
        For Each Site In TheExec.Sites
            g_alOTPDefValue = wOTPDefValue.Data
            Exit For
        Next Site
        
    
    Else
        ''''Here it's the 2nd Run and following Run
        For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
            g_OTPData.Category(lOtpIdx).lDefaultValue = g_OTPRev.Category(g_lRevIdx).DefaultValue(lOtpIdx)
            '20190521 Set init write value as default value
            ''''just update the activate sites, it's SiteVariant
            g_OTPData.Category(lOtpIdx).Write.Value = g_OTPRev.Category(g_lRevIdx).DefaultValue(lOtpIdx)
        Next lOtpIdx
        
'        For Each Site In TheExec.Sites
'            gD_wPGMData.Clear
'            gD_wReadData.Clear
'            gD_wPGMData.Data = g_alOTPDefValue
'        Next Site
    End If
    
    'JY 20200406 Need to do this every single touch down
    For Each Site In TheExec.Sites 'Don't use 'existing'
        gD_wPGMData.CreateConstant 0, CLng(g_iOTP_ADDR_TOTAL) * g_iOTP_REGDATA_BW, DspLong '20200402 move it to ResetDspGlobalVariable
        gD_wReadData.CreateConstant 0, CLng(g_iOTP_ADDR_TOTAL) * g_iOTP_REGDATA_BW, DspLong
        gD_wPGMData.Data = g_alOTPDefValue
    Next Site

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CreateDictionary_ForCalcSwCrc()
    Dim sFuncName As String: sFuncName = "CreateDictionary_ForCalcSwCrc"
    On Error GoTo ErrHandler
    Dim wsOTPTable As Worksheet
    Dim sKey As String
    Dim lRow As Long
    Dim lIdx  As Long
    Dim lNameCol  As Long
    '___Change "g_DicAHBCRC_Index" to "g_dictCRCByAHBRegName"
    '___No need to search in OTP_default_reglist, search OTP_Register_Map instead.
    g_dictCRCByAHBRegName.RemoveAll
    Set wsOTPTable = Application.ActiveWorkbook.Sheets(g_sOTP_SHEETNAME)
    'lRow = 1: lIdx = 0: lNameCol = 1
    lRow = 4: lIdx = 0: lNameCol = 4
    sKey = UCase(wsOTPTable.Cells(lRow, lNameCol))
    
    While sKey <> ""
        sKey = UCase(wsOTPTable.Cells(lRow, lNameCol))
        If g_dictCRCByAHBRegName.Exists(sKey) Then
        Else
            g_dictCRCByAHBRegName.Add sKey, lIdx
            lIdx = lIdx + 1
        End If
        
        lRow = lRow + 1: sKey = UCase(wsOTPTable.Cells(lRow, lNameCol))
    Wend
    If g_dictCRCByAHBRegName.Count = 0 Then TheExec.Datalog.WriteComment sFuncName & " Initail Fail, please check"
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'Public Sub OTPRead_VPPRampUp()
'    With TheHdw.DCVI(g_sVPP_PINNAME)
'        .Connect
'        .Gate = True
'        .SetCurrentAndRange 0.2, 0.2
'        .SetVoltageAndRange 1.5 * v, 10 * v
'    End With
'TheExec.Datalog.WriteComment "Power Up " & g_sVPP_PINNAME & " = " & Format(TheHdw.DCVI.Pins(g_sVPP_PINNAME).Voltage, "0.00") & "V"
'Exit Sub
'End Sub
Public Sub RampDownVpp_ForOtpRead() 'Down to 0V
    Dim sFuncName As String: sFuncName = "RampDownVpp_ForOtpRead"
    On Error GoTo ErrHandler
    Dim lVStep As Long
    
    With TheHdw.DCVI(g_sVPP_PINNAME)
        .Gate = False
        .Disconnect tlDCVIConnectDefault
        .Mode = tlDCVIModeVoltage
        .CurrentRange.AutoRange = True
        .SetCurrentAndRange 0.2, 0.2
        .SetVoltageAndRange 0#, 2# 'DC30 Vrange: 0.5/1/2/5/10/20/30
        .Gate = True
    End With
    
    For lVStep = 3 To 1 Step -1
        TheHdw.DCVI(g_sVPP_PINNAME).Voltage = 0.5 * lVStep
        TheHdw.Wait 0.01 * ms
    Next lVStep
    
    TheHdw.DCVI(g_sVPP_PINNAME).Gate = False
    TheHdw.DCVI(g_sVPP_PINNAME).Disconnect
    TheExec.Datalog.WriteComment "Power Down " & g_sVPP_PINNAME & " = " & Format(TheHdw.DCVI.Pins(g_sVPP_PINNAME).Voltage, "0.00") & "V"

Exit Sub
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub
Public Sub RampUpVpp_ForOtpBurn()
    Dim sFuncName As String: sFuncName = "RampUpVpp_ForOtpBurn"
    On Error GoTo ErrHandler
    Dim lVStep As Long

    With TheHdw.DCVI(g_sVPP_PINNAME)
        .Gate = False
        .Disconnect tlDCVIConnectDefault
        .Mode = tlDCVIModeVoltage
        .CurrentRange.AutoRange = True
        .SetCurrentAndRange 0.2, 0.2
        .SetVoltageAndRange 0#, 10# 'DC30 Vrange: 0.5/1/2/5/10/20/30
        .Connect
        .Gate = True
    End With
    
    For lVStep = 1 To 5 Step 1
        TheHdw.DCVI(g_sVPP_PINNAME).Voltage = 1.5 * lVStep
        TheHdw.Wait 0.01 * ms
    Next lVStep
    
    TheExec.Datalog.WriteComment "Power Up " & g_sVPP_PINNAME & " = " & Format(TheHdw.DCVI.Pins(g_sVPP_PINNAME).Voltage, "0.00") & "V"
    
Exit Sub
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub
Public Sub RampDownVpp_ForOtpBurn() 'Down to 0V
    Dim sFuncName As String: sFuncName = "RampDownVpp_ForOtpBurn"
    On Error GoTo ErrHandler
    Dim lVStep As Long

    For lVStep = 5 To 0 Step -1
        TheHdw.DCVI(g_sVPP_PINNAME).Voltage = 1.5 * lVStep
        TheHdw.Wait 0.1 * ms
    Next lVStep
    TheHdw.Wait 1 * ms
    With TheHdw.DCVI(g_sVPP_PINNAME)
        .Gate = False
        .Disconnect tlDCVIConnectDefault
    End With
    TheExec.Datalog.WriteComment "Power Down " & g_sVPP_PINNAME & " = " & Format(TheHdw.DCVI.Pins(g_sVPP_PINNAME).Voltage, "0.00") & "V"

Exit Sub
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'New function created to replace setreddecimal
'It updates the otpdata category Read section by input r_lOTPIdx
Public Function SetReadData2OTPCat_byOTPIdx(r_lOTPIdx As Long)
    Dim sFuncName As String: sFuncName = "SetReadData2OTPCat_byOTPIdx"
    On Error GoTo ErrHandler
    
    Dim wTemp As New DSPWave
    Dim lBitStart As Long
    Dim lBitEnd As Long
    Dim lBitWidth As Long
    Dim slReadData As New SiteLong
    Dim wParaReadData As New DSPWave
    't0 = theexec.Timer

    slReadData = 0
    lBitStart = g_OTPData.Category(r_lOTPIdx).lOtpBitStrStart
    lBitEnd = g_OTPData.Category(r_lOTPIdx).lOtpBitStrEnd
    lBitWidth = g_OTPData.Category(r_lOTPIdx).lBitWidth
    
    '___Get decimal value from gD_wReadData 20190515
    If TheExec.TesterMode = testModeOnline Then
    
        ''''20200313, need to check if it has benefit by RunDSP
        For Each Site In TheExec.Sites
            wTemp = gD_wReadData.Select(lBitStart, 1, lBitEnd - lBitStart + 1).Copy
            wParaReadData = wTemp.ConvertStreamTo(tldspParallel, lBitEnd - lBitStart + 1, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
            slReadData = wParaReadData.ElementLite(0)
        Next Site
    Else
        'Call rundsp.otp_get_ConvStream_ReadData(lBitStart, lBitEnd, slReadData)
        'Put the code to local reduce test time for 5ms to 3ms
        For Each Site In TheExec.Sites
            wTemp = gD_wPGMData.Select(lBitStart, 1, lBitEnd - lBitStart + 1).Copy
            wParaReadData = wTemp.ConvertStreamTo(tldspParallel, lBitEnd - lBitStart + 1, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
            slReadData = wParaReadData.Element(0)
        Next Site

    End If
    
    g_OTPData.Category(r_lOTPIdx).Read.Value = slReadData
    For Each Site In TheExec.Sites
        g_OTPData.Category(r_lOTPIdx).Read.HexStr(Site) = "0x" + Hex(slReadData(Site))
        g_OTPData.Category(r_lOTPIdx).Read.BitStrM(Site) = ConvertFormat_Dec2Bin_Complement(g_OTPData.Category(r_lOTPIdx).Read.Value(Site), lBitEnd - lBitStart + 1)
    Next Site
            
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'It updates the otpdata category Write section by input r_lOTPIdx
Public Function SetWriteData2OTPCat_byOTPIdx(r_lOTPIdx As Long)
    Dim sFuncName As String: sFuncName = "SetWriteData2OTPCat_byOTPIdx"
    On Error GoTo ErrHandler
    Dim lBitStart As Long
    Dim lBitEnd As Long
    Dim lBitWidth As Long
    Dim slWriteData As New SiteLong

    slWriteData = 0
    lBitStart = g_OTPData.Category(r_lOTPIdx).lOtpBitStrStart
    lBitEnd = g_OTPData.Category(r_lOTPIdx).lOtpBitStrEnd
    lBitWidth = g_OTPData.Category(r_lOTPIdx).lBitWidth
    
    '___Get decimal value from gD_wPGMData
    Call RunDsp.otp_get_ConvStream_WriteData(lBitStart, lBitEnd, slWriteData)
    g_OTPData.Category(r_lOTPIdx).Write.Value = slWriteData ''was CLng(slWriteData)
        
    For Each Site In TheExec.Sites
        g_OTPData.Category(r_lOTPIdx).Write.HexStr(Site) = "0x" + Hex(slWriteData(Site))
        g_OTPData.Category(r_lOTPIdx).Write.BitStrM(Site) = ConvertFormat_Dec2Bin_Complement(slWriteData(Site), lBitWidth)
    Next Site
            
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___[Notice] because it's uniform across Sites, so it's no issue w/o SiteLoop while using the DSPWave
'20190524 Init AHB info during OTP register map parsing
Public Function GetAhbMask(ByVal lFieldWidth As Long, ByVal lFieldOffset As Long, r_wRtnSerDataVal As DSPWave, r_lRtnParDataVal As Long) As Long
    Dim sFuncName As String: sFuncName = "GetAhbMask"
    On Error GoTo ErrHandler
    Dim wInitVal As New DSPWave
    Dim wRtnParDataVal As New DSPWave
    Dim IIdx As Integer
    Dim alBit() As Long

    wInitVal.CreateConstant 1, g_iAHB_BW, DspLong
        
    If (False) Then
        ''''20200313   '2020/03/28 SOMEHOW affect other blocks !!!
        alBit = wInitVal.Data
        For IIdx = lFieldOffset To lFieldOffset + lFieldWidth - 1
            alBit(IIdx) = 0
        Next
        wInitVal.Data = alBit
    Else
        For IIdx = lFieldOffset To lFieldOffset + lFieldWidth - 1
            wInitVal.Element(IIdx) = "0"
        Next
    End If
    
    ''''20200313, it's same size, so just Copy '2020/03/28 OK
    r_wRtnSerDataVal = wInitVal.Copy ''''wInitVal.Select(0, 1, g_iAHB_BW).Copy
    
    wRtnParDataVal = wInitVal.ConvertStreamTo(tldspParallel, g_iAHB_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
    r_lRtnParDataVal = wRtnParDataVal.Element(0)
         
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___20200313, it was called in CheckCrcConsistency_FW(), but CheckCrcConsistency_FW() does NOT be called.
'20190526 FW (New function in MP3P, need to debug)
Public Function CheckOtpWriteReadData()
    Dim sFuncName As String: sFuncName = "CheckOtpWriteReadData"
    On Error GoTo ErrHandler
    Dim lOtpIdx As Long, lOTPIdxCnt As Long: lOTPIdxCnt = (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
    
    Dim slWriteOTP() As New SiteLong, slReadOTP() As New SiteLong ', aslCalcReadAHBOTP() As New SiteLong
    Dim asOTPRegName() As String, sRegName() As String, asvCheckResults() As New SiteVariant, asBWOffset() As String
    Dim sOTPRegName As String, lAHBOffset As Long, lAHBBW As Long
    Dim asOtpOwner() As String, asDefaultReal() As String
    
    ReDim slWriteOTP(lOTPIdxCnt): ReDim slReadOTP(lOTPIdxCnt) ': ReDim aslCalcReadAHBOTP(lOTPIdxCnt)
    ReDim asOTPRegName(lOTPIdxCnt): ReDim sRegName(lOTPIdxCnt): ReDim asvCheckResults(lOTPIdxCnt): ReDim asBWOffset(lOTPIdxCnt)
    ReDim asOtpOwner(lOTPIdxCnt): ReDim asDefaultReal(lOTPIdxCnt) As String
    
    Dim sComment           As String

    Dim sbCheckResult  As New SiteBoolean
    
    If TheExec.TesterMode = testModeOffline Then
       TheExec.Datalog.WriteComment ""
       TheExec.Datalog.WriteComment "** Offline Mode **"
    End If
    
    If TheExec.Sites.Selected.Count = 0 Then Exit Function
    
    sbCheckResult = True


    For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
       
       With g_OTPData.Category(lOtpIdx)
           sOTPRegName = .sOtpRegisterName
           asOTPRegName(lOtpIdx) = .sOtpRegisterName
           sRegName(lOtpIdx) = .sRegisterName
           asOtpOwner(lOtpIdx) = .sOTPOwner
           asDefaultReal(lOtpIdx) = .sDefaultORReal
           lAHBBW = .lBitWidth
           lAHBOffset = .lOtpRegOfs
           asBWOffset(lOtpIdx) = lAHBBW
       End With

       '1).Get OTP Write data
       Call auto_OTPCategory_GetWriteDecimal(sOTPRegName, slWriteOTP(lOtpIdx))
       '2).Get OTP Read data
       Call auto_OTPCategory_GetReadDecimal(sOTPRegName, slReadOTP(lOtpIdx))
        
        Dim slCheckOTPData As New SiteLong: slCheckOTPData = slWriteOTP(lOtpIdx)
        Dim slCheckAHBData As New SiteLong
        
        If TheExec.TesterMode = testModeOnline Then
            slCheckAHBData = slReadOTP(lOtpIdx)
        Else
            slCheckAHBData = slWriteOTP(lOtpIdx)
        End If
       
        For Each Site In TheExec.Sites.Selected
            'asvCheckResults = slCheckOTPData(Site).Subtract(slCheckAHBData(Site))
            asvCheckResults(lOtpIdx) = slCheckOTPData.Subtract(slCheckAHBData)
            If asvCheckResults(lOtpIdx) <> 0 Then sbCheckResult(Site) = False
        Next Site
          
    Next lOtpIdx


    'B).DataLog
    Dim sTemp As String
    Dim slFailCnt As New SiteLong
                      
       TheExec.Datalog.WriteComment "<" + sFuncName + ">, InstanceName::" & TheExec.DataManager.InstanceName
       
       
        For Each Site In TheExec.Sites.Selected
            sTemp = "" 'FormatLog("Comment", -20)
            sTemp = " [Site" & CStr(Site) & "]"
            sTemp = sTemp & "," & FormatLog("OTP-REGName", -g_lOTPCateNameMaxLen) & "," & FormatLog("OTP-WriteTrimCode", -25)
            sTemp = sTemp & "," & FormatLog("OTP-ReadTrimCode", -25)
            sTemp = sTemp & "," & FormatLog("Check", -10)
            sTemp = sTemp & "," & FormatLog("OTPOWNER", -10)
            sTemp = sTemp & "," & FormatLog("Default|Real", -15)
            
            TheExec.Datalog.WriteComment ""
            TheExec.Datalog.WriteComment "*********************************************************************************************** " & _
                                       "*********************************************************************************************************"
            TheExec.Datalog.WriteComment sTemp
            TheExec.Datalog.WriteComment "*********************************************************************************************** " & _
                                                "*********************************************************************************************************"
            slFailCnt(Site) = 0
            If sbCheckResult(Site) = False Then
                For lOtpIdx = 0 To UBound(asOTPRegName)
                    If InStr(UCase(g_asBYPASS_AHBOTP_CHECK), UCase(asOTPRegName(lOtpIdx))) = 0 Then
                        sComment = ""
                    Else
                        sComment = ""
                    End If

                    sTemp = ""
                    sTemp = " [Site" & CStr(Site) & "]"
                    'sTemp = sTemp & "," & FormatLog(asOTPRegName(lOTPIdx), -g_lOTPCateNameMaxLen) & "," & FormatLog("&H" + Hex(slWriteOTP(lOTPIdx)), -12)
                    'sTemp = sTemp & "," & FormatLog(sRegName(lOTPIdx), -g_lOTPCateNameMaxLen - 10) & "," & FormatLog("&H" + Hex(slReadOTP(lOTPIdx)), -12)
                    sTemp = sTemp & "," & FormatLog(asOTPRegName(lOtpIdx), -g_lOTPCateNameMaxLen) & "," & FormatLog(slWriteOTP(lOtpIdx), -25)
                    sTemp = sTemp & "," & FormatLog(slReadOTP(lOtpIdx), -25)
                    sTemp = sTemp & "," & FormatLog(asvCheckResults(lOtpIdx), -10)
                    sTemp = sTemp & "," & FormatLog(asOtpOwner(lOtpIdx), -10)
                    sTemp = sTemp & "," & FormatLog(asDefaultReal(lOtpIdx), -15)
                    
                     
                    If asvCheckResults(lOtpIdx) <> 0 Then slFailCnt(Site) = slFailCnt(Site) + 1

                    'Datalog for check fail only
                    If asvCheckResults(lOtpIdx) <> 0 Then
                          TheExec.Datalog.WriteComment sTemp
                    End If
                    
                Next lOtpIdx
            End If
            TheExec.Datalog.WriteComment ""
        Next Site


        
        'DATALOG:
        Dim sParmName As String
        sParmName = "OTPExpected_vs_OTPActual_Check-slFailCnt"
        
        If g_bOtpEnable = True Then
             '2019-01-15: OTP Write vs OTP Read should match.
            TheExec.Flow.TestLimit slFailCnt, 0, 0, , , , unitNone, , TName:=sParmName
            'Call Acore.Utilities.testlimit_lng(sParmName, slFailCnt, 0, 0, "N", False)
        Else
            TheExec.Flow.TestLimit slFailCnt, 0, 0, , , , unitNone, , TName:=sParmName
            'Call Acore.Utilities.testlimit_lng(sParmName, slFailCnt, 0, 999, "N", False)
        End If


Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'Public Function OTP_Simulation_Stage(Optional Stage As String, Optional Disable_SiteNumber As Long)
'
'
'    If UCase(Stage) Like "*ALL*BLANK*" Then
'        Call Simulation_All_Blank
'        g_bBlankFlag = True
'
'    End If
'
'
'    If UCase(Stage) Like "*ALL*BURN*" Then
'        Call Simulation_All_BURN
'        g_bBlankFlag = False
'
'    End If
'
'
'    If UCase(Stage) Like "*FAKE*VALUE*" Then
'        Call Simulation_FAKE_VALUE(Disable_SiteNumber)
'        g_bBlankFlag = True
'
'    End If
'
'End Function
'
'Public Function Simulation_All_Blank()
'On Error GoTo ErrHandler
'    Dim sFuncName As String: sFuncName = "Simulation_All_Blank"
'    Dim i As Long
''   Purepose make all Read Category is blank
'        If g_bEnWrdOTPFTProg = True Then Call ReadWaferDataToCat
'
'        For Each Site In theexec.Sites.Existing
'            For i = 0 To (g_Total_OTP - 1) 'irnore ECID information
'                If g_bEnWrdOTPFTProg = True Then
'
'                        g_OTPData.Category(i).Read.HexStr = g_OTPData.Category(i).Write.HexStr
'                        g_OTPData.Category(i).Read.Value = g_OTPData.Category(i).Write.Value
'                Else
'                    g_OTPData.Category(i).Read.HexStr = ""
'                    g_OTPData.Category(i).Read.Value = 0
'
'                End If
'            Next i
'
'        Next Site
'
'Exit Function
'ErrHandler:
'    theexec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function Simulation_All_BURN()
'
'On Error GoTo ErrHandler
'    Dim sFuncName As String: sFuncName = "Simulation_All_BURN"
''   Purepose make all Read Category is blank
''
'        Dim i As Long
'        For Each Site In theexec.Sites.Existing
'            For i = 0 To (g_Total_OTP - 1) 'irnore ECID information
'
'                g_OTPData.Category(i).Read.BitStrM() = g_OTPData.Category(i).Write.BitStrM()
'                g_OTPData.Category(i).Read.HexStr = g_OTPData.Category(i).Write.HexStr
'                g_OTPData.Category(i).Read.Value = g_OTPData.Category(i).Write.Value
'
'            Next i
'
'        Next Site
'
'Exit Function
'
'ErrHandler:
'    theexec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'
'Public Function Simulation_FAKE_VALUE(Optional Disable_SiteNumber As Long = -1)
'
'On Error GoTo ErrHandler
'    Dim sFuncName As String: sFuncName = "Simulation_FAKE_VALUE"
''   Purepose make all Read Category is blank
'
'        Dim i As Long
'        Dim SimulationValue As New SiteLong
'        Dim shutdown_site As New SiteBoolean
'
'
'
'    sbSaveSiteStatus = theexec.Sites.Selected
'    PROGRAMSITE = False
'
'    If Disable_SiteNumber <> -1 Then
'        For Each Site In theexec.Sites
'                PROGRAMSITE = True
'                PROGRAMSITE(Disable_SiteNumber) = False 'According to site number to disable site
'        Next Site
'        theexec.Sites.Selected = PROGRAMSITE
'    End If
'
'
'
'
'        For Each Site In theexec.Sites.Selected
'            For i = 0 To (g_Total_OTP - 1) 'irnore ECID information
'                 If UCase(g_OTPData.Category(i).sOtpRegisterName) Like "*id*block*" Then
'                 'skip simulation ECID
'
'                 ElseIf UCase(g_OTPData.Category(i).sDefaultorReal) Like "*REAL*" Then
'                        SimulationValue = 2 ^ g_OTPData.Category(i).lBitWidth / 2
'                        Call auto_OTPCategory_SetWriteDecimal(g_OTPData.Category(i).sOtpRegisterName, SimulationValue)
'                 End If
'            Next i
'        Next Site
'
'
'Exit Function
'ErrHandler:
'    theexec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'___20200313, Need to check when in FW/Multi-Shot mode ??
Public Function CheckOtpBlank(Optional CalSumData As SiteLong) 'FWDebug
    Dim sFuncName As String: sFuncName = "CheckOtpBlank"
    On Error GoTo ErrHandler
    Dim psReadPat As New PatternSet
    Dim sParmName As String
    'Dim CalSumData As New SiteLong

    If TheExec.Sites.Active.Count = 0 Then Exit Function
    If g_bOTPOneShot = True Then
        '___OneShot Read Pattern Definition
        psReadPat.Value = "OTP_CRC_MANUAL_INC_READ_DEBUG"
        '___Read all burned values and calculate the summation
        '___The device is not blank if the sum is not equal to zero
        Call OTP_READREG_DSP_ALL(psReadPat.Value) 'One shot
        TheHdw.StartStopwatch 'Timer start
        RunDsp.SumUpCatReadData CalSumData
        Call OTP_SPT_D(" *** OTP FW Pattern , Exe Time =  ") 'Timer stop
        '___Datalog
        sParmName = "OTP_BlankCheck"
        TheExec.Flow.TestLimit CalSumData, 0, 0, TName:=sParmName
    Else 'FW or multi-shot
        ''''20200313, Need to Check ?? how to calc CalSumData
        'CalSumData = mSL_ARRAY_PROGRAMMED_0.Add(mSL_OTP_CHIPID).Add(mSL_OTP__CRC_0)
        '___Datalog
        sParmName = "OTP_BlankCheck"
        TheExec.Flow.TestLimit CalSumData, 0, 0, TName:=sParmName
    End If

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


''''20190603 trial, 20200313 update
'___20200313 from MP7P, could be faster, it's for SW CRC Calculation
Public Function CreateDictionary_Ahb2OtpIdxs() As Boolean
    '___Create dictionary according to OTP_Register_Map column-4 "reg_name" (g_dictAHBRegToOTPDataIdx)
    '___Please put in the function OnProgramStarted()
    On Error GoTo ErrHandler
    Dim sFuncName As String: sFuncName = "CreateDictionary_Ahb2OtpIdxs"

#If True Then
    ''''20200313 from MP7P (faster)
    Dim i As Long, k As Long
    Dim sKeyName As String
    ReDim g_asAhbRegNameInOtp(g_Total_OTP) ''''20200313 declaration an initial arrary
    k = 0
    g_dictAHBRegToOTPDataIdx.RemoveAll '20200323
    
    For i = 0 To (g_Total_OTP - 1)
        sKeyName = UCase(g_OTPData.Category(i).sRegisterName)
        If sKeyName <> "" Then
            If g_dictAHBRegToOTPDataIdx.Exists(sKeyName) = True Then
                g_dictAHBRegToOTPDataIdx(sKeyName) = g_dictAHBRegToOTPDataIdx.Item(sKeyName) & "," & i
            Else
                g_dictAHBRegToOTPDataIdx(sKeyName) = i
                
                '''20200313 add, replace g_dictCRCByAHBRegName for TTR
                g_asAhbRegNameInOtp(k) = sKeyName
                k = k + 1
            End If
        End If
    Next i
    
    ''''20200313 final the dimension
    ReDim Preserve g_asAhbRegNameInOtp(k - 1)
#Else
    ''''Original
    Dim lOtpIdx, lOTP2ndIdx As Long
    Dim l1stIdx As Long
    Dim l2ndIdx As Long
    Dim s1stAHBName As String
    Dim s2ndAHBName As String
    Dim sKeyName As String
    Dim vIdxCombine As Variant
    
    g_dictAHBRegToOTPDataIdx.RemoveAll
    
    For lOtpIdx = 0 To (g_Total_OTP - 1) 'was UBound(g_OTPData.Category) '___20200313, AHB New Method
        l1stIdx = g_OTPData.Category(lOtpIdx).lOtpIdx
        s1stAHBName = UCase(g_OTPData.Category(lOtpIdx).sRegisterName)
        vIdxCombine = l1stIdx
        For lOTP2ndIdx = lOtpIdx + 1 To UBound(g_OTPData.Category)
            s2ndAHBName = UCase(g_OTPData.Category(lOTP2ndIdx).sRegisterName)
            If s1stAHBName = s2ndAHBName Then
                l2ndIdx = g_OTPData.Category(lOTP2ndIdx).lOtpIdx
                vIdxCombine = vIdxCombine & "," & l2ndIdx
            End If
        Next lOTP2ndIdx
        
        '___Create g_dictAHBRegToOTPDataIdx
        sKeyName = UCase(s1stAHBName)
        If g_dictAHBRegToOTPDataIdx.Exists(sKeyName) = False Then
            g_dictAHBRegToOTPDataIdx.Add sKeyName, vIdxCombine
            vIdxCombine = ""
        End If
    Next lOtpIdx
#End If

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'20190611
Private Function Print2Csv(r_sFuncName As String)
    Dim sFuncName As String: sFuncName = "Print2Csv"
    On Error GoTo ErrHandler
    Dim vCurrFolder As Variant
    Dim sDirectoryName As String
    Dim lIdx As Long

    vCurrFolder = CurDir() & "\OTPDATA"
    sDirectoryName = vCurrFolder & "\" & "OTPData_WR_Check" & ".csv"

    Open sDirectoryName For Output As #1

    For lIdx = 0 To UBound(m_avOtpDataInfo)
        Print #1, CStr(m_avOtpDataInfo(lIdx))
    Next lIdx

    Close #1

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'20190611
'___20200313, because new AHB method involved, we need to be careful the real OTP category size.
Public Function Dump_OTPWRData_Check_Append(r_sFuncName As String, WriteOrReadREG As g_eRegWriteRead)
    Dim sFuncName As String: sFuncName = "Dump_OTPWRData_Check_Append"
    On Error GoTo ErrHandler
    Dim lSiteCnt As Long
    Dim lOtpIdx As Long
    Dim lWriteFlag As Long
    Dim sWRIndicator As String
    Dim vCurrFolder As Variant
    Dim sDirectoryName As String
    
    lSiteCnt = TheExec.Sites.Selected.Count
    If (WriteOrReadREG = eREGWRITE) Then
        lWriteFlag = 1
        sWRIndicator = "Write"
    ElseIf (WriteOrReadREG = eREGREAD) Then
        lWriteFlag = 0
        sWRIndicator = "Read"
    End If
    
    '___first time calls "Dump_OTPWRData_Check_Append"
    If g_lDebugDumpCnt = 0 Then
        Call MakeFolder("\OTPDATA")
        vCurrFolder = CurDir() & "\OTPDATA"
        sDirectoryName = vCurrFolder & "\" & "OTPData_WR_Check" & ".csv"

        If Dir(sDirectoryName) <> "" Then
            SetAttr sDirectoryName, vbNormal
            Kill (sDirectoryName)
        End If
        
        ReDim m_avOtpDataInfo((g_Total_OTP - 1) + 3)
        
            m_avOtpDataInfo(0) = "Register name, Default/Real, Index"
            m_avOtpDataInfo(1) = ",,"
            m_avOtpDataInfo(2) = ",,"
        
        'For lOTPIdx = 0 To UBound(g_OTPData.Category)
        For lOtpIdx = 0 To (g_Total_OTP - 1)
            m_avOtpDataInfo(lOtpIdx + 3) = CStr(g_OTPData.Category(lOtpIdx).sOtpRegisterName) & "," & CStr(g_OTPData.Category(lOtpIdx).sDefaultORReal) & "," & CStr(g_OTPData.Category(lOtpIdx).lOtpIdx)
        Next lOtpIdx
    End If
    
    '___string handling
    For Each Site In TheExec.Sites
        For lOtpIdx = 0 To (g_Total_OTP - 1)
            If lOtpIdx = 0 Then
                m_avOtpDataInfo(0) = m_avOtpDataInfo(0) & "," & r_sFuncName
            
                If lWriteFlag = 1 Then
                    m_avOtpDataInfo(1) = m_avOtpDataInfo(1) & "," & sWRIndicator
                ElseIf lWriteFlag = 0 Then
                    m_avOtpDataInfo(1) = m_avOtpDataInfo(1) & "," & sWRIndicator
                End If
            
                m_avOtpDataInfo(2) = m_avOtpDataInfo(2) & "," & CStr("Site" & CStr(Site))
            End If
            
            If lWriteFlag = 1 Then
            m_avOtpDataInfo(lOtpIdx + 3) = m_avOtpDataInfo(lOtpIdx + 3) & "," & CStr(g_OTPData.Category(lOtpIdx).Write.Value(Site))
            Else
            m_avOtpDataInfo(lOtpIdx + 3) = m_avOtpDataInfo(lOtpIdx + 3) & "," & CStr(g_OTPData.Category(lOtpIdx).Read.Value(Site))
            End If
        Next lOtpIdx
    Next Site
            
    '___Print data to CSV
    Call Print2Csv(r_sFuncName)
    
    
    '___Record how many time the dump function was called
    g_lDebugDumpCnt = g_lDebugDumpCnt + 1
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___20200313, add for AHB New Method by OTPData Structure (copy from MP7P)
'___It is used to add other AHB Register into OTP data structure
'___could be considered to put in the function InitializeOtpDataFromAhbMap(), check it later
'Public Function init_parse_AHB_Table()
Public Function add_AHBReg_into_OTPData()
On Error GoTo ErrHandler
Dim sFuncName As String: sFuncName = "add_AHBReg_into_OTPData" ''was "init_parse_AHB_Table"

    Dim m_wsAHBTable As Worksheet
    Dim lRegIdx As Long
    Dim lOtpIdx As Long
    Dim MaxArrSize As Long
    Dim m_searchRegName As String
    Dim m_searchFiledName As String
    
    Dim lRegNameCol As Long
    Dim lRegAddrCol As Long
    Dim lRegIdxCol As Long
    Dim lFieldWidthCol As Long
    Dim lFieldOffsetCol As Long
    Dim lLastRow As Long
    
    Dim j, k As Integer
    Dim AHB_FieldWidth As Long
    Dim AHB_FieldOffset As Long
    Dim rtnSerDataValWave As New DSPWave
    Dim rtnParDataVal As Long     'update name, 20200313
    Dim temp_Cal_Deci_AHB_ByMaskOfs As New SiteLong
    
    Dim m_obj As Variant
    Dim mS_AHBReg_otpIdx As String
    Dim m_idxStr_arr As Variant

    MaxArrSize = g_Total_OTP + g_Total_AHB - g_OTP_With_AHB - 1
    
    If (IsSheetExists(g_sAHB_SHEETNAME) = False) Then
        GoTo ErrHandler
    Else
        '___Define AHB_register_map sheet name
        Set m_wsAHBTable = Sheets(g_sAHB_SHEETNAME)
    End If
    
    Call LocateColnRow_Ahb(lRegNameCol, lRegAddrCol, lRegIdxCol, lFieldWidthCol, lFieldOffsetCol, lLastRow)
    ''Call AHB_Rng_Def(col_reg_name, col_reg_addr, col_reg_idx, col_field_width, col_field_ofs, Reg_last_Row)
    
    lOtpIdx = g_Total_OTP - 1
    
    For lRegIdx = 2 To lLastRow
        m_searchRegName = UCase(CStr(m_wsAHBTable.Cells(lRegIdx, lRegNameCol).Value))
        m_searchFiledName = UCase(CStr(m_wsAHBTable.Cells(lRegIdx, lRegNameCol + 1).Value))
    
        If Not g_dictAHBRegToOTPDataIdx.Exists(m_searchRegName) Then  'check if OTPData have same reg name
            lOtpIdx = lOtpIdx + 1
            'm_wsAHBTable.Cells(lRegIdx, lRegNameCol).Interior.ColorIndex = 3
            With g_OTPData.Category(lOtpIdx)
                .sRegisterName = m_searchRegName
                .sName = m_searchFiledName
                .sAhbAddress = Replace(CStr(m_wsAHBTable.Cells(lRegIdx, lRegAddrCol).Value), "0x", "&H")
                
                 AHB_FieldWidth = CLng(m_wsAHBTable.Cells(lRegIdx, lFieldWidthCol).Value)
                 AHB_FieldOffset = CLng(m_wsAHBTable.Cells(lRegIdx, lFieldOffsetCol).Value)
                .lBitWidth = AHB_FieldWidth
                .lOtpRegOfs = AHB_FieldOffset
                .sOtpRegisterName = "NA"
                .lOtpIdx = lOtpIdx

                ''''20200313, do a trial inside, check the TT
                Call GetAhbMask(AHB_FieldWidth, AHB_FieldOffset, rtnSerDataValWave, rtnParDataVal) ''''was AHBMask_Calc()
            
                .lAhbMaskVal = rtnParDataVal

                For j = 0 To g_iAHB_BW - 1
                    .sAhbMask(j) = (rtnSerDataValWave.Data(j))
                Next j
                temp_Cal_Deci_AHB_ByMaskOfs = (rtnParDataVal) Xor (2 ^ g_iAHB_BW - 1)
                .lCalDeciAhbByMaskOfs = temp_Cal_Deci_AHB_ByMaskOfs.ShiftRight(AHB_FieldOffset)
                
            End With

            AHB_FieldOffset = 0
            AHB_FieldWidth = 0
    
        Else 'If rege exist, check all filed exist in g_OTPData.category
            If (0) Then 'Just for debug
                mS_AHBReg_otpIdx = g_dictAHBRegToOTPDataIdx.Item(m_searchRegName)
                m_idxStr_arr = Split(mS_AHBReg_otpIdx, ",")
                For k = 0 To UBound(m_idxStr_arr)
                    With g_OTPData.Category(m_idxStr_arr(k))
                        If .sName = m_searchFiledName Then
                         Exit For
                            If k = UBound(m_idxStr_arr) Then Stop 'Should not stop
                        End If
                    End With
                Next
            End If
        End If
    Next lRegIdx
        
''EndParsing:
    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function








