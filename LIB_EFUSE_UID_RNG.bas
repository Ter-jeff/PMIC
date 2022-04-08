Attribute VB_Name = "LIB_EFUSE_UID_RNG"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As LongPtr, pszContainer As String, pszProvider As String, ByVal dwProvType As LongPtr, ByVal dwFlags As LongPtr) As Long
    Private Declare PtrSafe Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwLen As LongPtr, pbBuffer As Byte) As Long
    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwFlags As LongPtr) As Long
'    Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, pszContainer As String, pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
'    Private Declare PtrSafe Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, pbBuffer As Byte) As Long
'    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
#Else
'    Private Declare Function CryptAcquireContext Lib "C:\windows\system32\advapi32_AES.dll" Alias "CryptAcquireContextA" (phProv As Long, pszContainer As String, pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
'    Private Declare Function CryptGenRandom Lib "C:\windows\system32\advapi32_AES.dll" (ByVal hProv As Long, ByVal dwLen As Long, pbBuffer As Byte) As Long
'    Private Declare Function CryptReleaseContext Lib "C:\windows\system32\advapi32_AES.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
#End If

Private Const CRYPT_VERIFYCONTEXT   As Long = &HF0000000
Private Const PROV_RSA_FULL         As Long = 1

Private Const RNG_BUFFER_SIZE As Long = 4096

Private Type RNG_RandomBuffer_type
  Buffer(1 To RNG_BUFFER_SIZE) As Byte
  pos As Long
End Type

Private RNG_RandomBuffer As RNG_RandomBuffer_type

Public Function RNG_cryptoRandomByte() As Byte

On Error GoTo errHandler
    Dim funcName As String:: funcName = "RNG_cryptoRandomByte"
    
    Dim RandomVal As Byte
    Dim Error As Boolean
    Dim i As Long
    Dim Sum As Double
 
    If RNG_RandomBuffer.pos = 0 Then
        ' either this is first run, assuming global variables are zero initilised
        ' or the buffer has wrapped
        
        ' refill the buffer
        Error = RNG_GetRandom()
        If (Error <> True) Then
            MsgBox ("RNG_cryptoRandomByte: Error getting random buffer")
            End
        End If
    
        ' perform a basic sanity test that the buffer isn't full of zeros
            Sum = 0
        For i = 1 To RNG_BUFFER_SIZE
            Sum = Sum + RNG_RandomBuffer.Buffer(i)
        Next i
        If (Sum = 0) Then
            MsgBox ("RNG_cryptoRandomByte: Random buffer is zero")
            End
        End If
    End If
    
    ' plus one since the array starts at index 1
    RandomVal = RNG_RandomBuffer.Buffer(RNG_RandomBuffer.pos + 1)
    
    ' progress the buffer read position
    RNG_RandomBuffer.pos = (RNG_RandomBuffer.pos + 1) Mod RNG_BUFFER_SIZE
    
    RNG_cryptoRandomByte = RandomVal

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Private Function RNG_GetRandom() As Boolean

On Error GoTo errHandler
    Dim funcName As String:: funcName = "RNG_GetRandom"

    Dim phProv As LongPtr
    Dim result1 As Long
    Dim result2 As Long
    Dim result3 As Long
    Dim String1 As String, MsgStr1 As String
    Dim String2 As String, MsgStr2 As String
    Dim String3 As String, MsgStr3 As String
 
 
    ' If TheExec.TesterMode = testModeOnline Then
        result1 = CryptAcquireContext(phProv, ByVal vbNullChar, ByVal vbNullChar, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
        String1 = "CryptAcquireContext32"
    ' Else
    '    result1 = CryptAcquireContext64(phProv, ByVal vbNullChar, ByVal vbNullChar, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
    '    String1 = "CryptAcquireContext64"
    ' End If
 
    If (result1 <> 1) Then
        MsgStr1 = "RNG_GetRandom: Call to " & String1 & " FAILED with error: 0x" & "Hex(Err.LastDllError)"
        'MsgBox ("RNG_GetRandom: Call to CryptAcquireContext FAILED with error: 0x" & Hex(Err.LastDllError))
        MsgBox MsgStr1
        'End
    End If

    'If TheExec.TesterMode = testModeOnline Then
        result2 = CryptGenRandom(phProv, RNG_BUFFER_SIZE, RNG_RandomBuffer.Buffer(1))
        String2 = "CryptAcquireContext32"
    'Else
    '    result2 = CryptGenRandom64(phProv, RNG_BUFFER_SIZE, RNG_RandomBuffer.Buffer(1))
    '    String2 = "CryptAcquireContext64"
    'End If

    If (result2 <> 1) Then
        MsgStr2 = "RNG_GetRandom: Call to " & String2 & " FAILED with error: 0x" & "Hex(Err.LastDllError)"
        'MsgBox ("RNG_GetRandom: Call to CryptGenRandom FAILED with error: 0x" & Hex(Err.LastDllError))
        MsgBox MsgStr2
        'End
    End If
   
    ' If TheExec.TesterMode = testModeOnline Then
        result3 = CryptReleaseContext(phProv, 0)
        String3 = "CryptReleaseContext32"
    ' Else
    '    result3 = CryptReleaseContext64(phProv, 0)
    '    String3 = "CryptReleaseContext64"
    ' End If
 

    If (result3 <> 1) Then
        MsgStr3 = "RNG_GetRandom: Call to " & String3 & " FAILED with error: 0x" & "Hex(Err.LastDllError)"
        'MsgBox ("RNG_GetRandom: Call to CryptReleaseContext FAILED with error: 0x" & Hex(Err.LastDllError))
        MsgBox MsgStr3
        'End
    End If
   
    RNG_GetRandom = result1 And result2 And result3

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
