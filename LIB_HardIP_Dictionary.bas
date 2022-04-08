Attribute VB_Name = "LIB_HardIP_Dictionary"
Option Explicit
Public gDictDebug As New Dictionary                                                         'MTR Record
'**************************************************
'SeaHawk Edited by 20190606
Public gl_SpecialString As String
Public gl_DictString As New Dictionary
Public gl_DictDSPWave As New Dictionary
'**************************************************
Public gDictDSPWaves As New Dictionary
Private gDictCurrMeasurements As New Dictionary
Private gDictSiteLong As New Dictionary
Private RegDict As New Dictionary

' Function to retrieve a current/voltage/frequency measurement. This
' function can be called by user interpose functions to access
' previously stored measurements
Public Function GetStoredMeasurement(KeyName As String) As Variant
    KeyName = LCase(KeyName)
    If Not gDictCurrMeasurements.Exists(KeyName) Then
        TheExec.ErrorLogMessage "Stored measurement " & KeyName & " not found."
    Else
        Set GetStoredMeasurement = gDictCurrMeasurements(KeyName)
    End If
End Function

' Function to store a measurement for later retrieval, typically from a custom user
' postMeasure interpose function
Public Function AddStoredMeasurement(KeyName As String, ByRef obj As Variant)
    KeyName = LCase(KeyName)
    If gDictCurrMeasurements.Exists(KeyName) Then
        gDictCurrMeasurements.Remove (KeyName)
    End If
    gDictCurrMeasurements.Add KeyName, obj
End Function

' Function to retrieve a captured DSPWave. This
' function can be called by user interpose functions to access
' previously stored data
Public Function GetStoredCaptureData(KeyName As String) As Variant
    KeyName = LCase(KeyName)
    If Not gDictDSPWaves.Exists(KeyName) Then
        TheExec.ErrorLogMessage "Stored capture data " & KeyName & " not found."
    Else
        Set GetStoredCaptureData = gDictDSPWaves(KeyName)
    End If
End Function

' Function to store a measurement for later retrieval, typically from a custom user
' postMeasure interpose function
Public Function AddStoredCaptureData(KeyName As String, ByRef obj As DSPWave)
    KeyName = LCase(KeyName)
    If gDictDSPWaves.Exists(KeyName) Then
        gDictDSPWaves.Remove (KeyName)
    End If
    gDictDSPWaves.Add KeyName, obj
End Function

Public Function RemoveAllStored()
    gDictCurrMeasurements.RemoveAll
    gDictDSPWaves.RemoveAll
    gDictSiteLong.RemoveAll
End Function


Public Function AddStoredData(KeyName As String, obj As SiteDouble)
    KeyName = LCase(KeyName)
    If gDictSiteLong.Exists(KeyName) Then
        gDictSiteLong.Remove (KeyName)
    End If
    gDictSiteLong.Add KeyName, obj

End Function
Public Function GetStoredData(KeyName As String) As Variant
    KeyName = LCase(KeyName)
    If Not gDictSiteLong.Exists(KeyName) Then
        TheExec.ErrorLogMessage "Stored capture data " & KeyName & " not found."
    Else
        Set GetStoredData = gDictSiteLong(KeyName)
    End If
End Function
Public Function StoredRegAssign()
Dim i As Long
RegDict.RemoveAll
For i = 0 To UBound(RegAssignInfo.ByTest)
    If RegDict.Exists(RegAssignInfo.ByTest(i).testName & "_" & "ModeA") Or RegDict.Exists(RegAssignInfo.ByTest(i).testName & "_" & "ModeB") Then
    TheExec.ErrorLogMessage "Duplicate Assign on" & RegAssignInfo.ByTest(i).testName
    Else
    RegDict.Add RegAssignInfo.ByTest(i).testName & "_" & "ModeA", RegAssignInfo.ByTest(i).RtnByModeA
    RegDict.Add RegAssignInfo.ByTest(i).testName & "_" & "ModeB", RegAssignInfo.ByTest(i).RtnByModeB
    End If
Next i

End Function

Public Function GetRegFromDictByTestByMode(RegAssignment As String, RegAssignChecker As Boolean) As String

If RegDict.Exists(RegAssignment) Then
    RegAssignChecker = True
    RegAssignment = RegDict(RegAssignment)
Else
    RegAssignChecker = False
    'TheExec.ErrorLogMessage ("Your RegAssignment is not specified in Reg_Assign Sheet")
    'Debug.Print ("Your RegAssignment is not specified in Reg_Assign Sheet")
End If
End Function
Public Function Public_AddStoredCaptureData(KeyName As String, ByRef obj As DSPWave) As Long
'**************************************************
'SeaHawk Edited by 20190606
'**************************************************
    KeyName = LCase(KeyName)
    If gl_DictDSPWave.Exists(KeyName) Then
        gl_DictDSPWave.Remove (KeyName)
    End If
    gl_DictDSPWave.Add KeyName, obj
End Function

Public Function Public_GetStoredCaptureData(KeyName As String) As Variant
'**************************************************
'SeaHawk Edited by 20190606
'**************************************************
    KeyName = LCase(KeyName)
    If Not gl_DictDSPWave.Exists(KeyName) Then
        TheExec.ErrorLogMessage "Stored capture data " & KeyName & " not found."
    Else
        Set Public_GetStoredCaptureData = gl_DictDSPWave(KeyName)
    End If
End Function

Public Function Public_AddStoredString(KeyName As String, ByRef obj As String)
'**************************************************
'SeaHawk Edited by 20190606
'**************************************************
    KeyName = LCase(KeyName)
    If gl_DictString.Exists(KeyName) Then
        gl_DictString.Remove (KeyName)
    End If
    gl_DictString.Add KeyName, obj
End Function


''20190604AddFunction
Public Function IsExists_StoredCaptureData(KeyName As String) As Boolean
    KeyName = LCase(KeyName)
    IsExists_StoredCaptureData = gDictDSPWaves.Exists(KeyName)
End Function
Public Function Public_GetStoredString(KeyName As String) As Variant
'**************************************************
'SeaHawk Edited by 20190606
'**************************************************
    KeyName = LCase(KeyName)
    If Not gl_DictString.Exists(KeyName) Then
        TheExec.ErrorLogMessage "Stored measurement " & KeyName & " not found."
    Else
'        Set Public_GetStoredString = gl_DictString(KeyName)
        gl_SpecialString = gl_DictString(KeyName)
    End If
End Function
