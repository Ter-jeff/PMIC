Attribute VB_Name = "LIB_EFUSE_UID_CRC"
Option Explicit
Option Base 0
Public gL_CRCidx As Long

Public Sub CRC_ComputeCRCforBit(ByRef CRC() As Byte, bit As Byte)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CRC_ComputeCRCforBit"
    
    Dim inv As Byte
    
    inv = bit Xor CRC(31)
    CRC(31) = CRC(30)
    CRC(30) = CRC(29)
    CRC(29) = CRC(28)
    CRC(28) = CRC(27)
    CRC(27) = CRC(26)
    CRC(26) = CRC(25) Xor inv
    CRC(25) = CRC(24)
    CRC(24) = CRC(23)
    CRC(23) = CRC(22) Xor inv
    CRC(22) = CRC(21) Xor inv
    CRC(21) = CRC(20)
    CRC(20) = CRC(19)
    CRC(19) = CRC(18)
    CRC(18) = CRC(17)
    CRC(17) = CRC(16)
    CRC(16) = CRC(15) Xor inv
    CRC(15) = CRC(14)
    CRC(14) = CRC(13)
    CRC(13) = CRC(12)
    CRC(12) = CRC(11) Xor inv
    CRC(11) = CRC(10) Xor inv
    CRC(10) = CRC(9) Xor inv
    CRC(9) = CRC(8)
    CRC(8) = CRC(7) Xor inv
    CRC(7) = CRC(6) Xor inv
    CRC(6) = CRC(5)
    CRC(5) = CRC(4) Xor inv
    CRC(4) = CRC(3) Xor inv
    CRC(3) = CRC(2)
    CRC(2) = CRC(1) Xor inv
    CRC(1) = CRC(0) Xor inv
    CRC(0) = inv

Exit Sub

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next

End Sub

Public Sub CRC_Zero_Array(ByRef CRC() As Byte)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CRC_Zero_Array"

    Dim i As Long
    
    For i = 0 To UBound(CRC)
        CRC(i) = 0
    Next

Exit Sub

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
    
End Sub

Public Sub CRC16_ComputeCRCforBit(ByRef CRC() As Byte, bit As Byte, Optional DebugInfo = False)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CRC16_ComputeCRCforBit"
    
    Dim inv As Byte
    Dim i As Long
    
    Dim TempStr As String

    ''''20161026 update
    If (gL_CRCidx = 0 And DebugInfo) Then
        TempStr = "====== CRC shift serial ========"
        TheExec.Datalog.WriteComment (TempStr)
    End If
    TempStr = FormatNumeric("CRC shift (" & FormatNumeric(gL_CRCidx, 4) & ") : ", 17)

    inv = bit Xor CRC(15)
    CRC(15) = CRC(14)
    CRC(14) = CRC(13)
    CRC(13) = CRC(12) Xor inv
    CRC(12) = CRC(11) Xor inv
    CRC(11) = CRC(10) Xor inv
    CRC(10) = CRC(9) Xor inv
    CRC(9) = CRC(8)
    CRC(8) = CRC(7) Xor inv
    CRC(7) = CRC(6)
    CRC(6) = CRC(5) Xor inv
    CRC(5) = CRC(4) Xor inv
    CRC(4) = CRC(3)
    CRC(3) = CRC(2)
    CRC(2) = CRC(1) Xor inv
    CRC(1) = CRC(0)
    CRC(0) = inv

    For i = 0 To 15
        TempStr = TempStr & CRC(i)
    Next i
    
    If DebugInfo Then
        TheExec.Datalog.WriteComment (TempStr) + "[LSB...MSB]"
    End If
    gL_CRCidx = gL_CRCidx + 1

Exit Sub

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next

End Sub
