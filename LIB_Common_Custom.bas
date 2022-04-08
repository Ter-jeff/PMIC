Attribute VB_Name = "LIB_Common_Custom"
Option Explicit
'Revision History:
'V0.0 initial bring up

'variable declaration
Public Const Version_Lib_Common_Custom = "0.1"  'lib version

Public Function GroupPinsByMod(MyArray() As String, mod_count As Integer) As Variant()


Dim i, j, X As Integer


'Dim MyArray() As Variant
'
'MyArray = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o")

Dim MyArrayCount As Integer

MyArrayCount = UBound(MyArray) + 1


Dim GroupNumber As Integer
Dim GroupNumberMod As Integer

GroupNumber = MyArrayCount \ mod_count
GroupNumberMod = MyArrayCount Mod mod_count

Dim GroupPin() As Variant
ReDim GroupPin(GroupNumber)

For i = 0 To GroupNumber - 1
    
    GroupPin(i) = MyArray(i * mod_count)
    
            For j = 1 To mod_count - 1
                GroupPin(i) = GroupPin(i) & "," & MyArray(i * mod_count + j)
            Next j
    
    
'    Select Case mod_count
'        Case 1:
'            GroupPin(i) = MyArray(i)
'        Case 2:
'            GroupPin(i) = MyArray(i * mod_count) & "," & MyArray(i * mod_count + 1)
'        Case 3:
'            GroupPin(i) = MyArray(i * mod_count) & "," & MyArray(i * mod_count + 1) & "," & MyArray(i * mod_count + 2)
'        Case 4:
'            GroupPin(i) = MyArray(i * mod_count) & "," & MyArray(i * mod_count + 1) & "," & MyArray(i * mod_count + 2) & "," & MyArray(i * mod_count + 3)
'        Case 5:
'            GroupPin(i) = MyArray(i * mod_count) & "," & MyArray(i * mod_count + 1) & "," & MyArray(i * mod_count + 2) & "," & MyArray(i * mod_count + 3) & "," & MyArray(i * mod_count + 4)
'    End Select


Next i

    If GroupNumberMod = 0 Then
    
            ReDim Preserve GroupPin(GroupNumber - 1)
    
    Else
            GroupPin(i) = ""
    
            GroupPin(i) = MyArray(GroupNumber * mod_count)
    
            For j = 1 To GroupNumberMod - 1 Step 1
                GroupPin(i) = GroupPin(i) & "," & MyArray(GroupNumber * mod_count + j)
            Next j
    
    End If


'    Select Case GroupNumberMod
'        Case 1:
'            GroupPin(i) = MyArray(GroupNumber * mod_count)
'        Case 2:
''            GroupPin(i) = MyArray(GroupNumber * mod_count) & "," & MyArray(GroupNumber * mod_count + 1)
'                For j = 1 To GroupNumberMod - 1 Step 1
'                    GroupPin(i) = GroupPin(i) & "," & MyArray(GroupNumber * mod_count + j)
'                Next j
'        Case 3:
'            GroupPin(i) = MyArray(GroupNumber * mod_count) & "," & MyArray(GroupNumber * mod_count + 1) & "," & MyArray(GroupNumber * mod_count + 2)
'        Case 4:
'            GroupPin(i) = MyArray(GroupNumber * mod_count) & "," & MyArray(GroupNumber * mod_count + 1) & "," & MyArray(GroupNumber * mod_count + 2) & "," & MyArray(GroupNumber * mod_count + 3)
'        Case 5:
'            GroupPin(i) = MyArray(GroupNumber * mod_count) & "," & MyArray(GroupNumber * mod_count + 1) & "," & MyArray(GroupNumber * mod_count + 2) & "," & MyArray(GroupNumber * mod_count + 3) & "," & MyArray(GroupNumber * mod_count + 4)
'    End Select

'thehdw.Wait 0.01

    GroupPinsByMod = GroupPin()


End Function

Function max(lng1 As Double, lng2 As Double) As Double
    max = lng1
    If lng2 >= lng1 Then max = lng2
End Function

Function Min(lng1 As Double, lng2 As Double) As Double
    Min = lng1
    If lng2 <= lng1 Then Min = lng2
End Function


Sub sbHideASheet(Sheet As String)
    Sheets(Sheet).Visible = False
End Sub

Sub sbUnhideASheet(Sheet As String)
        '$$ManifestSheet
    Sheets(Sheet).Visible = True
End Sub

Function RepeatChr(Str As String, repeat As Long) As String
    Dim i As Long
    RepeatChr = ""
    For i = 0 To repeat - 1
        RepeatChr = RepeatChr & Str
    Next i
End Function

Public Function FormatNumericDatalog(num As Variant, length As Long, LeftZero As Boolean) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "FormatNumericDatalog"
    
    ''''Example
    ''''----------------------------------------
    '''' length > 0  is to right shift
    '''' length < 0  is to left  shift
    ''''----------------------------------------
    ''''FormatNumeric(123456, 8) + "...end"
    ''''  123456...end
    ''''
    ''''FormatNumeric(123456,-8) + "...end"
    ''''123456  ...end
    ''''
    ''''----------------------------------------
    
    Dim numStr As String
    Dim tmpLen As Long
    Dim spcLen As Long
    
    numStr = CStr(num)
    tmpLen = Len(numStr)
    
    If (tmpLen > Abs(length)) Then
        spcLen = 0
    Else
        spcLen = Abs(length) - tmpLen
    End If
    
    If (length < 0) Then   ''''number shift to the very left
        FormatNumericDatalog = CStr(num) + Space(spcLen)
    ElseIf LeftZero Then ''''default: shift to the very right
        FormatNumericDatalog = Space(spcLen) + CStr(num)
    Else
        FormatNumericDatalog = CStr(num) + Space(spcLen)
    End If
    
Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function
