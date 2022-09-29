Attribute VB_Name = "VBT_LIB_Spotcal"
Option Explicit
'Revision History:
'V0.0 initial

Global glb_spotcalval As New Dictionary

Public gI_DelCheck_idx As Integer
'''''''''''''''''''''''''''''
'Definition of glb_spotcalval
'itemName = High_Side_Pin + "_" + Low_Side_Pin + "_" + Vrange
'itemVal = site0_val + "_" + site1_val + "_" + site2_val ...
'''''''''''''''''''''''''''''

Function spotcal_Pre_OnProgramValidated() As Long
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "spotcal_Pre_OnProgramValidated"

    Dim pin_H     As New PinList
    Dim pin_L     As New PinList

    'clear all before re-cal
    glb_spotcalval.RemoveAll

    ''    'buck2p LP PRdson
    ''    pin_H.Value = "ATB0_UVI80_DM"
    ''    pin_L.Value = "ATB3_UVI80_DM"
    ''    Call runSpotcal(pin_H, pin_L, 1.4 * v)
    ''
    ''    pin_H.Value = "VDDC_UVI80_DM"
    ''    pin_L.Value = "ATB2_UVI80_DM"
    ''    Call runSpotcal(pin_H, pin_L, 1.4 * v)
    ''
    ''    pin_H.Value = "VDDC_UVI80_DM"
    ''    pin_L.Value = "ATB1_UVI80_DM"
    ''    Call runSpotcal(pin_H, pin_L, 1.4 * v)

    Exit Function
    
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Private Sub runSpotcal(High_Side_Pin As PinList, _
                       Low_Side_Pin As PinList, _
                       Vrange As Double)
    On Error GoTo ErrorHandler
    Dim funcName  As String:: funcName = "runSpotcal"

    Dim measLoDiscon As New SiteDouble
    Dim measHiDiscon As New SiteDouble
    Dim compVal   As New SiteDouble
    Dim ItemName  As String
    Dim itemStr   As String
    Dim OriginalSiteStatus As New SiteBoolean
    Dim AllTrueSiteStatus As New SiteBoolean

    ''''''''    OriginalSiteStatus = TheExec.Sites.Selected
    ''''''''    AllTrueSiteStatus = True
    ''''''''    TheExec.Sites.Selected = AllTrueSiteStatus


    'Initial Setup
    With TheHdw.DCDiffMeter.Pins(High_Side_Pin)
        .MeterMode = tlDCDiffMeterModeHighAccuracy
        .VoltageRange = Vrange * V
        .HardwareAverage = 64
    End With

    'High_Side short, Low_Side disconnect Measurement
    With TheHdw.DCDiffMeter.Pins(High_Side_Pin)
        .LowSide.Pins = High_Side_Pin
        .Connect
    End With
    measLoDiscon = TheHdw.DCDiffMeter.Pins(High_Side_Pin).Read(tlStrobe, 1, 100 * KHz, tlDCDiffMeterReadingFormatAverage)
    TheHdw.DCDiffMeter.Pins(High_Side_Pin).Disconnect

    'Low_Side short, High_Side disconnect Measurement
    With TheHdw.DCDiffMeter.Pins(Low_Side_Pin)
        .LowSide.Pins = (Low_Side_Pin)
        .Connect
    End With
    measHiDiscon = TheHdw.DCDiffMeter.Pins(Low_Side_Pin).Read(tlStrobe, 1, 100 * KHz, tlDCDiffMeterReadingFormatAverage)

    With TheHdw.DCDiffMeter.Pins(Low_Side_Pin)
        .Disconnect
        .MeterMode = tlDCDiffMeterModeHighSpeed    'Generally use HighSpeed mode with 8.33MHz sample rate
    End With
    'Calculation
    compVal = measLoDiscon.Add(measHiDiscon).Divide(2)

    'Offline simulation
    If TheExec.TesterMode = testModeOffline Then
        For Each Site In TheExec.Sites.Existing
            compVal(Int(Site)) = Rnd()
        Next Site
    End If

    'Update dictionary
    'if there's site shut down, we will have less site compval and cause trouble during getSpotcalVal
    If TheExec.Sites.Active.Count < TheExec.Sites.Existing.Count Then
        ''''''''    If TheExec.Sites.Active.Count <> TheExec.Sites.Selected.Count Then
        Debug.Print "There's site shut down before Spotcal for " + ItemName + "!!!"
        'Calc and add to dictionary
        ''----------------------------------------------------------------
        'this code for save sitedouble value into dictionary
        ItemName = High_Side_Pin.Value + "_" + Low_Side_Pin.Value + "_" + CStr(Vrange)
        glb_spotcalval.Add ItemName, compVal
        ''----------------------------------------------------------------
    Else
        ''--------------------------------------------------------------
        '        For Each Site In TheExec.Sites.Selected
        '            ItemName = High_Side_Pin.Value + "_" + Low_Side_Pin.Value + "_" + CStr(Vrange) + "_" + CStr(Site)
        '            itemStr = CStr(compVal(Int(Site)))
        '            glb_spotcalval.Add ItemName, itemStr
        '        Next Site
        ''----------------------------------------------------------------
        'this code for save sitedouble value into dictionary
        ItemName = High_Side_Pin.Value + "_" + Low_Side_Pin.Value + "_" + CStr(Vrange)
        glb_spotcalval.Add ItemName, compVal
        ''----------------------------------------------------------------

    End If

    ''''''''    '' restore the Site status to original setting
    ''''''''    TheExec.Sites.Selected = OriginalSiteStatus
    '20181001
    getSpotcalVal High_Side_Pin, Low_Side_Pin, Vrange

    Exit Sub
ErrorHandler:
    'LIB_ErrorDescription ("runSpotcal")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Private Function doSpotcal(High_Side_Pin As PinList, _
                           Low_Side_Pin As PinList, _
                           Vrange As Double) As SiteDouble
    On Error GoTo ErrorHandler
    Dim funcName  As String:: funcName = "doSpotcal"

    Dim measLoDiscon As New SiteDouble
    Dim measHiDiscon As New SiteDouble
    Dim compVal   As New SiteDouble
    Dim ItemName  As String
    Dim itemStr   As String


    Dim OriginalSiteStatus As New SiteBoolean
    Dim AllTrueSiteStatus As New SiteBoolean
    OriginalSiteStatus = TheExec.Sites.Selected
    AllTrueSiteStatus = True
    TheExec.Sites.Selected = AllTrueSiteStatus




    'Initial Setup
    With TheHdw.DCDiffMeter.Pins(High_Side_Pin)
        .MeterMode = tlDCDiffMeterModeHighAccuracy
        .VoltageRange = Vrange * V
        .HardwareAverage = 64
    End With

    TheHdw.DCDiffMeter.Pins(High_Side_Pin).Disconnect
    TheHdw.DCDiffMeter.Pins(Low_Side_Pin).Disconnect

    'High_Side short, Low_Side disconnect Measurement
    With TheHdw.DCDiffMeter.Pins(High_Side_Pin)
        .LowSide.Pins = High_Side_Pin
        .Connect
    End With
    measLoDiscon = TheHdw.DCDiffMeter.Pins(High_Side_Pin).Read(tlStrobe, 1, 100 * KHz, tlDCDiffMeterReadingFormatAverage)
    TheHdw.DCDiffMeter.Pins(High_Side_Pin).Disconnect

    'Low_Side short, High_Side disconnect Measurement
    With TheHdw.DCDiffMeter.Pins(Low_Side_Pin)
        .LowSide.Pins = (Low_Side_Pin)
        .Connect
    End With
    measHiDiscon = TheHdw.DCDiffMeter.Pins(Low_Side_Pin).Read(tlStrobe, 1, 100 * KHz, tlDCDiffMeterReadingFormatAverage)
    TheHdw.DCDiffMeter.Pins(Low_Side_Pin).Disconnect
    'Calculation
    compVal = measLoDiscon.Add(measHiDiscon).Divide(2)

    'Offline simulation
    If TheExec.TesterMode = testModeOffline Then
        For Each Site In TheExec.Sites.Existing
            compVal(Int(Site)) = Rnd()
        Next Site
    End If

    '' restore the Site status to original setting
    TheExec.Sites.Selected = OriginalSiteStatus


    ''-----------------------------------------------------------------------------------
    ''    ItemName = High_Side_Pin.Value + "_" + Low_Side_Pin.Value + "_" + CStr(Vrange) + "_" + CStr(TheExec.Sites.SiteNumber)
    ''    itemStr = CStr(compVal(Int(TheExec.Sites.SiteNumber)))
    ''    glb_spotcalval.Add ItemName, itemStr
    ''doSpotcal = compVal(Int(TheExec.Sites.SiteNumber))
    ''-----------------------------------------------------------------------------------
    ItemName = High_Side_Pin.Value + "_" + Low_Side_Pin.Value + "_" + CStr(Vrange)
    glb_spotcalval.Add ItemName, compVal
    Set doSpotcal = compVal
    ''-----------------------------------------------------------------------------------






    Exit Function
ErrorHandler:
    'LIB_ErrorDescription ("DoSpotCal")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function getSpotcalVal(High_Side_Pin As PinList, _
                              Low_Side_Pin As PinList, _
                              Vrange As Double) As SiteDouble

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "getSpotcalVal"

    Dim ItemName  As String
    Dim tempVal   As New SiteDouble
    Dim siteCompStr As String
    Dim arrCompStr() As String


    ''''--------------------------------------------------------------------------------------------------------
    ''    For Each Site In TheExec.Sites.Selected
    ''        ItemName = High_Side_Pin.Value + "_" + Low_Side_Pin.Value + "_" + CStr(Vrange) + "_" + CStr(Site)
    ''        If glb_spotcalval.Exists(ItemName) Then
    ''            siteCompStr = glb_spotcalval.Item(ItemName)
    ''        Else
    ''            siteCompStr = doSpotcal(High_Side_Pin, Low_Side_Pin, Vrange)
    ''        End If
    ''
    ''        If Len(trim(siteCompStr)) = 0 Then
    ''            If glb_spotcalval.Exists(ItemName) Then glb_spotcalval.Remove (ItemName)
    ''            siteCompStr = doSpotcal(High_Side_Pin, Low_Side_Pin, Vrange)
    ''        End If
    ''
    ''        tempVal = CDbl(siteCompStr)
    ''    Next Site
    '''------------------------------------------------------------------------------------------------------------
    ''---- this code for sitedouble value output
    ItemName = High_Side_Pin.Value + "_" + Low_Side_Pin.Value + "_" + CStr(Vrange)
    If glb_spotcalval.Exists(ItemName) Then
        tempVal = glb_spotcalval.Item(ItemName)
    Else
        tempVal = doSpotcal(High_Side_Pin, Low_Side_Pin, Vrange)
    End If
    ''''--------------------------------------------------------------------------------------------------------




    '''    ' ---- below is open SpotCal.txt and write spot information to file.
    '''    If TheExec.Datalog.Setup.LotSetup.TestMode = gtlTestMode.gtl_Engineeringmode Then
    '''        Call ExportSpotCalCommand(CStr(High_Side_Pin), CStr(Low_Side_Pin), Vrange)
    '''        '20181001
    '''        '=================================================
    '''        Dim fs As New FileSystemObject
    '''        Dim St_WriteTxtFile As TextStream
    '''        mS_File = ".\REGCHECK\SpotCal.txt"
    '''        Set St_WriteTxtFile = fs.OpenTextFile(mS_File, ForAppending, True)
    '''
    '''
    '''            For Each Site In TheExec.Sites.Selected
    '''                ItemName = High_Side_Pin.Value + "_" + Low_Side_Pin.Value + "_" + CStr(Vrange) + "_" + CStr(Site)
    '''                St_WriteTxtFile.WriteLine ItemName & "=" & tempVal(Site)
    '''            Next Site
    '''            St_WriteTxtFile.WriteLine ("")
    '''            St_WriteTxtFile.WriteLine Now
    '''            St_WriteTxtFile.WriteLine ("======================================================")
    '''        '=================================================
    '''    End If

    Set getSpotcalVal = tempVal

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Private Sub ExportSpotCalCommand(High_Side_Pin As String, _
'                                 Low_Side_Pin As String, _
'                                 Vrange As Double)
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "ExportSpotCalCommand"
'
'    Dim mS_FILETYPE As String
'    Dim mS_File   As String
'    Dim fs        As New FileSystemObject
'    Dim St_WriteTxtFile As TextStream
'    Dim objSheet  As Worksheet
'
'
'    '20181009 Delete file
'
'    If gI_DelCheck_idx = 0 Then
'        If IsFileExists(".\REGCHECK\SpotCal.txt") = True Then
'            fs.DeleteFile ".\REGCHECK\SpotCal.txt"
'        End If
'        gI_DelCheck_idx = gI_DelCheck_idx + 1
'    End If
'
'    mS_File = ".\REGCHECK\SpotCal.txt"
'
'    Set St_WriteTxtFile = fs.OpenTextFile(mS_File, ForAppending, True)
'    St_WriteTxtFile.WriteLine ("pin_H.Value = " & """" & High_Side_Pin & """")
'    St_WriteTxtFile.WriteLine ("pin_L.Value = " & """" & Low_Side_Pin & """")
'    St_WriteTxtFile.WriteLine ("Call runSpotcal(pin_H, pin_L, " & CStr(Vrange) & ")")
'    St_WriteTxtFile.WriteLine ("")
'
'    Exit Sub
'ErrHandler:
'    RunTimeError funcName
'    If AbortTest Then Exit Sub Else Resume Next
'End Sub

'20181009
Private Function IsFileExists(ByVal strFileName As String) As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IsFileExists"


    If Dir(strFileName, 16) <> Empty Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'Private Sub InitSpotCalFile()
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "InitSpotCalFile"
'
'    Dim mS_FILETYPE As String
'    Dim mS_File   As String
'    mS_File = ".\REGCHECK\SpotCal.txt"
'    Call File_CheckAndCreateFolder(".\REGCHECK\")
'    Call File_CreateAFile(mS_File, "")
'
'
'    Exit Sub
'ErrHandler:
'    RunTimeError funcName
'    If AbortTest Then Exit Sub Else Resume Next
'End Sub




