VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibTestNameGrp7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private arrGenTName() As String
Private Const Grp8forCode = "C"
Private Const Grp8forT = "T"
Private Const priDefault = "X"
Private Const priNoTrim = "P"
Private Const priGrpValue = "EMPTY"
Private priTarget As String
Private priUnit   As String

Private Enum TNameGroup
    Grp5 = 4
    Grp6 = 5
    Grp7 = 6
    Grp8 = 7
    Grp9 = 8
    Grp10 = 9
    Grp11 = 10
    Grp12 = 11
End Enum

Public Function SetGenTNameArray(arrTName() As String)
    ReDim arrGenTName(UBound(arrTName))
    arrGenTName = arrTName
End Function

Public Function GetGenTNameArray() As String()
    GetGenTNameArray = arrGenTName
End Function

Public Function SetGenTNameTarget(Target As String)
    priTarget = CorrectTestName(Target)
End Function

Public Function SetGenTNameUnit(Unit As String)
    priUnit = Unit
End Function

Public Function GetGenTNameUnit() As String
    GetGenTNameUnit = priUnit
End Function

Public Function PreTrim() As String
    arrGenTName(TNameGroup.Grp7) = "PreTrim"
    arrGenTName(TNameGroup.Grp8) = Grp8forT
    PreTrim = Join(arrGenTName, "_")
End Function

Public Function PreTrimCode() As String
    arrGenTName(TNameGroup.Grp7) = "PreTrimCode"
    arrGenTName(TNameGroup.Grp8) = Grp8forCode
    PreTrimCode = Join(arrGenTName, "_")
End Function

Public Function PreTrimDelta() As String
    arrGenTName(TNameGroup.Grp7) = "PreTrimDelta"
    arrGenTName(TNameGroup.Grp8) = priDefault
    PreTrimDelta = Join(arrGenTName, "_")
End Function

Public Function PostTrim(Optional Pilot As Boolean = False) As String
    arrGenTName(TNameGroup.Grp7) = "PostTrim"

    If Pilot = True Then arrGenTName(TNameGroup.Grp8) = Grp8forT

    PostTrim = Join(arrGenTName, "_")
End Function

Public Function PostTrimCode(Optional Pilot As Boolean = False) As String
    arrGenTName(TNameGroup.Grp7) = "PostTrimCode"

    If Pilot = True Then arrGenTName(TNameGroup.Grp8) = Grp8forCode

    PostTrimCode = Join(arrGenTName, "_")
End Function


Public Function PostTrimDelta() As String
    arrGenTName(TNameGroup.Grp7) = "PostTrimDelta"
    arrGenTName(TNameGroup.Grp8) = priDefault
    PostTrimDelta = Join(arrGenTName, "_")
End Function

'''Public Function TweakTrim() As String
'''    arrGenTName(TNameGroup.Grp7) = "TweakTrim"
'''    arrGenTName(TNameGroup.Grp8) = Grp8forT
'''    TweakTrim = Join(arrGenTName, "_")
'''End Function
'''
'''Public Function TweakTrimCode() As String
'''    arrGenTName(TNameGroup.Grp7) = "TweakTrimCode"
'''    arrGenTName(TNameGroup.Grp8) = Grp8forCode
'''    TweakTrimCode = Join(arrGenTName, "_")
'''End Function

Public Function FinalTrim() As String
    arrGenTName(TNameGroup.Grp7) = "FinalTrim"
    arrGenTName(TNameGroup.Grp8) = Grp8forT
    arrGenTName(TNameGroup.Grp12) = priTarget & priUnit
    FinalTrim = Join(arrGenTName, "_")
End Function

Public Function FinalTrimCode() As String
    arrGenTName(TNameGroup.Grp7) = "FinalTrimCode"
    arrGenTName(TNameGroup.Grp8) = Grp8forCode
    FinalTrimCode = Join(arrGenTName, "_")
End Function

Public Function FinalTrimDelta() As String
    arrGenTName(TNameGroup.Grp7) = "FinalTrimDelta"
    arrGenTName(TNameGroup.Grp8) = priDefault
    FinalTrimDelta = Join(arrGenTName, "_")
End Function

Public Function PostBurn() As String
    arrGenTName(TNameGroup.Grp7) = "PostBurn"
    arrGenTName(TNameGroup.Grp8) = Grp8forT
    arrGenTName(TNameGroup.Grp12) = priTarget & priUnit
    PostBurn = Join(arrGenTName, "_")
End Function

Public Function PostBurnCode() As String
    arrGenTName(TNameGroup.Grp7) = "PostBurnCode"
    arrGenTName(TNameGroup.Grp8) = Grp8forCode
    PostBurnCode = Join(arrGenTName, "_")
End Function

Public Function PostBurnDelta() As String
    arrGenTName(TNameGroup.Grp7) = "PostBurnDelta"
    arrGenTName(TNameGroup.Grp8) = priDefault
    PostBurnDelta = Join(arrGenTName, "_")
End Function

Public Function TrimLink(Code As Variant) As String
    If Code <> "" Then
        arrGenTName(TNameGroup.Grp7) = "TrimLink"
        arrGenTName(TNameGroup.Grp8) = CorrectTestName(Format(CStr(Code), "000"))
    Else
        arrGenTName(TNameGroup.Grp7) = priDefault
    End If
    TrimLink = Join(arrGenTName, "_")
End Function

Public Function TrimTarget() As String
    arrGenTName(TNameGroup.Grp7) = "TrimTarget"
    arrGenTName(TNameGroup.Grp8) = Grp8forT
    TrimTarget = Join(arrGenTName, "_")
End Function

Public Function PostBurnTarget() As String
    arrGenTName(TNameGroup.Grp7) = "PostBurnTarget"
    arrGenTName(TNameGroup.Grp8) = Grp8forT
    PostBurnTarget = Join(arrGenTName, "_")
End Function

Public Function PostBurnGNG(LoLimit As String, HiLimit As String) As String
    arrGenTName(TNameGroup.Grp7) = "PostBurnGNG"
    If LoLimit <> "" Then
        arrGenTName(TNameGroup.Grp11) = CorrectTestName(LoLimit)
    Else
        arrGenTName(TNameGroup.Grp11) = priDefault
    End If
    If LoLimit <> "" Then
        arrGenTName(TNameGroup.Grp12) = CorrectTestName(HiLimit)
    Else
        arrGenTName(TNameGroup.Grp12) = priDefault
    End If

    PostBurnGNG = Join(arrGenTName, "_")
End Function

Public Function PostBurnGNGCHAR(LoLimit As String, HiLimit As String) As String
    arrGenTName(TNameGroup.Grp7) = "PostBurnGNGCHAR"
    If LoLimit <> "" Then
        arrGenTName(TNameGroup.Grp11) = CorrectTestName(LoLimit)
    Else
        arrGenTName(TNameGroup.Grp11) = priDefault
    End If
    If LoLimit <> "" Then
        arrGenTName(TNameGroup.Grp12) = CorrectTestName(HiLimit)
    Else
        arrGenTName(TNameGroup.Grp12) = priDefault
    End If
    PostBurnGNGCHAR = Join(arrGenTName, "_")
End Function

'''Public Function Addr() As String
'''    arrGenTName(TNameGroup.Grp7) = "Addr"
'''    Addr = Join(arrGenTName, "_")
'''End Function

Public Function Default() As String
    arrGenTName(TNameGroup.Grp7) = priDefault
    Default = Join(arrGenTName, "_")
End Function

'Public Function NoTrim(Optional Group5 As String = priGrpValue) As String
'    Dim tmpGrp5   As String
'    tmpGrp5 = arrGenTName(TNameGroup.Grp5)
'    If Group5 <> priGrpValue Then arrGenTName(TNameGroup.Grp5) = Group5
'    arrGenTName(TNameGroup.Grp6) = priDefault
'    arrGenTName(TNameGroup.Grp7) = priNoTrim
'    NoTrim = Join(arrGenTName, "_")
'    arrGenTName(TNameGroup.Grp5) = tmpGrp5
'End Function

Private Function CorrectTestName(GrpName As String) As String
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Customized"

    If (InStr(GrpName, "_") = 1) Then
        GrpName = Replace(GrpName, "_", "", 1, 1, vbTextCompare)
    End If

    If InStr(GrpName, "_") > 1 Then
        GrpName = Replace(GrpName, "_", "-", vbTextCompare)
    End If

    If (Mid(GrpName, Len(GrpName), 1) = "-") Then
        GrpName = Mid(GrpName, 1, Len(GrpName) - 1)
    End If

    CorrectTestName = GrpName

    Exit Function
ErrHandler:
    RunTimeError funcName
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

