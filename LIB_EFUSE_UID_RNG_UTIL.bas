Attribute VB_Name = "LIB_EFUSE_UID_RNG_UTIL"
Option Explicit


Public Const RNG_UTIL_BYTE_SIZE As Long = 8

' fill_array_with_byte_from_str(
'                                                   binaryStr    - string of "1"s and "0"s, of size RNG_UTIL_BYTE_SIZE
'                                                   toFill       - string array to write into
'                                                   startPos     - where in toFill to start writing
'                                               )
' returns the position in the array after the last filled position
'
Public Function RNG_UTIL_fill_array_with_byte_from_str(binaryStr As String, ByRef toFill() As String, startPos As Long) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "RNG_UTIL_fill_array_with_byte_from_str"

    Dim i As Long
    Dim valueToFill As Long
    
    If Len(binaryStr) <> RNG_UTIL_BYTE_SIZE Then
        MsgBox ("RNG_UTIL_fill_array_with_byte_from_str: binaryStr isn't RNG_UTIL_BYTE_SIZE")
        End
    End If
    
    If (startPos + RNG_UTIL_BYTE_SIZE - 1) > UBound(toFill) Then
        MsgBox ("RNG_UTIL_fill_array_with_byte_from_str: can't write past end of toFill")
        End
    End If
    
    For i = 1 To RNG_UTIL_BYTE_SIZE
        valueToFill = CInt(Mid(binaryStr, i, 1))
        
        If (valueToFill < 0) Or (valueToFill > 1) Then
            MsgBox ("RNG_UTIL_fill_array_with_byte_from_str: binaryStr contains something other than 0 or 1")
            End
        End If
        
        toFill(startPos + i - 1) = CStr(valueToFill)
    Next i
    
    RNG_UTIL_fill_array_with_byte_from_str = startPos + RNG_UTIL_BYTE_SIZE

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' byte_to_str (
'                      inputByte    -   byte to convert
'              )
' returns the binary representation of inputByte as string of "1"s and "0"s
'
Public Function RNG_UTIL_byte_to_str(inputByte As Byte) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "RNG_UTIL_byte_to_str"

    Dim X As Byte
    Dim y As Byte
    Dim output As String
    
    X = inputByte
    
    While X > 0
        y = Fix(X / 2)
        
        If y = X / 2 Then
            output = "0" & output
        Else
            output = "1" & output
        End If
        X = y
    Wend
    
    While Len(output) < RNG_UTIL_BYTE_SIZE
        output = "0" & output
    Wend
    
    RNG_UTIL_byte_to_str = output
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
