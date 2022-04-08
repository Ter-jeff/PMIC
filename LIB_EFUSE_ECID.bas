Attribute VB_Name = "LIB_EFUSE_ECID"
Option Explicit

''''---------------------------------------------------------------------------------------------------
'''' ECID Fuse
''''---------------------------------------------------------------------------------------------------

''''20160414, was auto_Hex2Binary and rename to auto_LotIDCh2Binary
Public Function auto_LotIDCh2Binary(InputStr As String) As String
                                                                                                                         
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_LotIDCh2Binary"
    
    Dim i As Long
    Dim PerChar As String
    Dim PerLetter As String
    Dim BinStr As String
    Dim j As Long
    Dim DecodeBin As String
    Dim MyArray() As Variant
    Dim myArrayBin() As Variant

    InputStr = UCase(InputStr)

    MyArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                      "A", "B", "C", "D", "E", "F", "G", "H", "I", _
                      "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                      "S", "T", "U", "V", "W", "X", "Y", "Z")

    myArrayBin = Array("000000", "000001", "000010", "000011", "000100", "000101", "000110", "000111", "001000", "001001", _
                    "001010", "001011", "001100", "001101", "001110", "001111", "010000", "010001", "010010", _
                    "010011", "010100", "010101", "010110", "010111", "011000", "011001", "011010", "011011", _
                    "011100", "011101", "011110", "011111", "100000", "100001", "100010", "100011")

    BinStr = ""

    For i = 1 To Len(InputStr)
        PerChar = Mid(InputStr, i, 1)
        'One-to-One mapping, myarray() mappping to myarraybin()
        For j = 0 To UBound(MyArray)
            If (PerChar = MyArray(j)) Then
               DecodeBin = myArrayBin(j)
               Exit For
            End If
        Next j
        BinStr = BinStr + DecodeBin
    Next i
    auto_LotIDCh2Binary = BinStr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function auto_Binary2Hex(InputStr As String) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Binary2Hex"
    
    Dim i As Long
    Dim PerChar As String
    Dim PerLetter As String
    Dim BinStr As String
    Dim j As Long
    Dim HexChar As String
    Dim MyArray() As Variant
    Dim myArrayBin() As Variant
    Dim strlen As String
    Dim cnt As Long
                                                                                                                             
    MyArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                       "A", "B", "C", "D", "E", "F")
    myArrayBin = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", "1000", "1001", _
                     "1010", "1011", "1100", "1101", "1110", "1111")

    BinStr = ""
    cnt = 0
    strlen = Len(InputStr)
                                                                                                                             
    For i = 1 To strlen Step 4
        cnt = cnt + 1
        PerChar = Mid(InputStr, (strlen + 1) - cnt * 4, 4)
        For j = 0 To UBound(myArrayBin)
            If (PerChar = myArrayBin(j)) Then
               HexChar = MyArray(j)
               Exit For
            End If
        Next j
                                                                                                                         
        BinStr = HexChar + BinStr
        'Debug.Print "Binstr = " & BinStr & ", i=" & i
    Next i
                                                                                                                         
    auto_Binary2Hex = BinStr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_WaferData_to_HexECID(m_lotid As Variant, WaferID As Variant, Xcoor As Variant, Ycoor As Variant) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_WaferData_to_HexECID"
    
    Dim BinStr As String
    Dim m_binarr() As Long
    
    '< Step1. transfer lotid, waferid, x and y coord to binary string>
    BinStr = ""
    BinStr = BinStr + auto_LotIDCh2Binary(CStr(m_lotid))
    BinStr = BinStr + auto_Dec2Bin_EFuse(WaferID, WAFERID_BITWIDTH, m_binarr)
    BinStr = BinStr + auto_Dec2Bin_EFuse(Xcoor, XCOORD_BITWIDTH, m_binarr)
    BinStr = BinStr + auto_Dec2Bin_EFuse(Ycoor, YCOORD_BITWIDTH, m_binarr)
    
    Do Until (Len(BinStr) Mod 64 = 0)
        BinStr = BinStr + "0"
    Loop
    
    '< Step2. Convert the joint binary string to user Hexadecimal code >
    auto_WaferData_to_HexECID = auto_Binary2Hex(StrReverse(BinStr))

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_MappingCharToBinStr(InputStr As String) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MappingCharToBinStr"
    
    Dim i As Long
    Dim MyArray() As Variant, myArrayBin() As Variant
    '=== Initialization ===
    'for LotId

    MyArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                   "A", "B", "C", "D", "E", "F", "G", "H", "I", _
                   "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                   "S", "T", "U", "V", "W", "X", "Y", "Z")
    myArrayBin = Array("000000", "000001", "000010", "000011", "000100", "000101", "000110", "000111", "001000", "001001", _
                 "001010", "001011", "001100", "001101", "001110", "001111", "010000", "010001", "010010", _
                 "010011", "010100", "010101", "010110", "010111", "011000", "011001", "011010", "011011", _
                 "011100", "011101", "011110", "011111", "100000", "100001", "100010", "100011")
                          

    For i = 0 To UBound(MyArray)
      If UCase(InputStr) = MyArray(i) Then
            auto_MappingCharToBinStr = myArrayBin(i)
            Exit For
      End If

    Next i

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function auto_MappingBinStrtoChar(InputStr As String) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MappingBinStrtoChar"
    
    Dim i As Long
    Dim MyArray() As Variant, myArrayBin() As Variant
    '=== Initialization ===
    'for LotId

    MyArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                   "A", "B", "C", "D", "E", "F", "G", "H", "I", _
                   "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                   "S", "T", "U", "V", "W", "X", "Y", "Z")
    myArrayBin = Array("000000", "000001", "000010", "000011", "000100", "000101", "000110", "000111", "001000", "001001", _
                 "001010", "001011", "001100", "001101", "001110", "001111", "010000", "010001", "010010", _
                 "010011", "010100", "010101", "010110", "010111", "011000", "011001", "011010", "011011", _
                 "011100", "011101", "011110", "011111", "100000", "100001", "100010", "100011")
                          

    For i = 0 To UBound(MyArray)
        If InputStr Like myArrayBin(i) Then
            auto_MappingBinStrtoChar = MyArray(i)
            Exit For
        Else
            auto_MappingBinStrtoChar = "?" ''''201808XX add
        End If
    Next i

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function

''''This function will convert the decimal to binary
''''<Notice> Means that MSB is at BitArray(0).
Public Function auto_Dec2Bin(ByVal n As Long, ByRef BinArray() As Long)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Dec2Bin"
    
    Dim i As Long, j As Long
    Dim Element_Amount As Long
    Dim Count As Long
    '               01101
    ' BinArray(4) 1 (LSB)
    ' BinArray(3) 0
    ' BinArray(2) 1
    ' BinArray(1) 1
    ' BinArray(0) 0 (MSB)

    'Example:: n=11 (MSB.01011.LSB)
    'By this function Dec2Bin, bit0 is MSB
    'bit=0, value= 0 (MSB)
    'bit=1, value= 1
    'bit=2, value= 0
    'bit=3, value= 1
    'bit=4, value= 1 (LSB)


    Element_Amount = UBound(BinArray)
    If n > (2 ^ (Element_Amount + 1) - 1) Then
        TheExec.Datalog.WriteComment "Error(auto_Dec2Bin): Overange for " & n
        n = 0
        GoTo errHandler ''''20170715 update
    End If

    For j = 0 To Element_Amount
        BinArray(j) = 0
    Next j

    If n < 0 Then MsgBox ("Warning(auto_Dec2Bin)!!! Decimal Number should be positive integer")
    i = 0
    Do Until n = 0
        If (i > Element_Amount) Then TheExec.Datalog.WriteComment "Warning (auto_Dec2Bin)!!! Decimal " & n & " is over-range (>" & i & "bit)"
        If (n Mod 2) Then
            BinArray(Element_Amount - i) = 1
        Else
            BinArray(Element_Amount - i) = 0
        End If
        n = Int(n / 2)
        i = i + 1
    Loop

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


''''eFuse_Pgm_Bit will be (MSBFirst=N), here eFuse_Pgm_Bit(start_bit)=LSB
''''eFuse_Pgm_Bit(0) = LSB
''''bin_str = [MSB...LSB]
''''20170331 update ByRef start_bit to ByVal start_bit
Public Function auto_CalcEfuseBit(idx As Long, DecVal As Long, ByVal start_bit As Long, length As Long, ByRef eFuse_Pgm_Bit() As Long, ByRef bin_str As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CalcEfuseBit"
    
    Dim BinArray() As Long
    ReDim BinArray(length - 1)
    Dim j As Long
    Dim m_bitsum As Long

    Dim ss As Variant
    ss = TheExec.sites.SiteNumber

    Call auto_Dec2Bin(DecVal, BinArray) ''''MSB is at BinArray(0)
    bin_str = ""
    m_bitsum = 0
    For j = length - 1 To 0 Step -1 'To length - 1
        eFuse_Pgm_Bit(start_bit) = BinArray(j) ''''here j=length - 1 is LSB
        bin_str = CStr(BinArray(j)) + bin_str
        start_bit = start_bit + 1
        m_bitsum = m_bitsum + BinArray(j)
    Next j

    ''''20150601 New
    ECIDFuse.Category(idx).Write.Decimal(ss) = CLng(DecVal)
    ECIDFuse.Category(idx).Write.Value(ss) = DecVal
    ECIDFuse.Category(idx).Write.ValStr(ss) = CStr(DecVal)
    ECIDFuse.Category(idx).Write.BitStrM(ss) = bin_str
    ECIDFuse.Category(idx).Write.BitStrL(ss) = StrReverse(bin_str)
    ECIDFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
    ECIDFuse.Category(idx).Write.HexStr(ss) = "0x" & auto_BinStr2HexStr(bin_str, CLng(Ceiling(length / 4)))  ''''20161018 add 20170331 update  "0x"

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''eFuse_Pgm_Bit will be (MSBFirst=Y), here eFuse_Pgm_Bit(start_bit)=MSB
''''eFuse_Pgm_Bit(0) = MSB
''''bin_str = [MSB...LSB]
''''20170331 update ByRef start_bit to ByVal start_bit
Public Function auto_CalcEfuseBit_Reversed(idx As Long, DecVal As Long, ByVal start_bit As Long, length As Long, ByRef eFuse_Pgm_Bit() As Long, ByRef bin_str As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CalcEfuseBit_Reversed"

    Dim BinArray() As Long
    ReDim BinArray(length - 1)
    Dim j As Long
    Dim m_bitsum As Long

    Dim ss As Variant
    ss = TheExec.sites.SiteNumber

    Call auto_Dec2Bin(DecVal, BinArray) ''''MSB is at BinArray(0)
    bin_str = ""
    m_bitsum = 0
    For j = 0 To length - 1
        eFuse_Pgm_Bit(start_bit) = BinArray(j)  ''''here j=0 is MSB
        bin_str = bin_str + CStr(BinArray(j))
        start_bit = start_bit + 1
        m_bitsum = m_bitsum + BinArray(j)
    Next j

    ''''20150601 New
    ECIDFuse.Category(idx).Write.Decimal(ss) = CLng(DecVal)
    ECIDFuse.Category(idx).Write.Value(ss) = DecVal
    ECIDFuse.Category(idx).Write.ValStr(ss) = CStr(DecVal)
    ECIDFuse.Category(idx).Write.BitStrM(ss) = bin_str
    ECIDFuse.Category(idx).Write.BitStrL(ss) = StrReverse(bin_str)
    ECIDFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
    ECIDFuse.Category(idx).Write.HexStr(ss) = "0x" & auto_BinStr2HexStr(bin_str, CLng(Ceiling(length / 4))) ''''20161018 add  20170331 update  "0x"

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_Decompose_StrArray_to_BitArray(ByVal FuseType As String, SingleStrArray() As String, SingleBitArray() As Long, singleBitSum As Long)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Decompose_StrArray_to_BitArray"
    
    Dim i As Long, j As Long, tmpStr As String
    Dim cnt As Long
    
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber
    
    Dim BitsPerRow As Long
    Dim ReadCycles As Long
    Dim BitsPerCycle As Long
    Dim BitsPerBlock As Long

    FuseType = UCase(Trim(FuseType))
    If (FuseType = "ECID") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = EcidBitsPerRow             ''''32   , 16  , 32
        ReadCycles = EcidReadCycle              ''''16   , 16  , 16
        BitsPerCycle = ECIDBitPerCycle          ''''32   , 32  , 32
        BitsPerBlock = EcidBitPerBlockUsed      ''''256  , 256 , 512
        
    ElseIf (FuseType = "CFG") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = EConfigBitsPerRow          ''''32   , 16  , 32
        ReadCycles = EConfigReadCycle           ''''32   , 32  , 16
        BitsPerCycle = EConfigReadBitWidth      ''''32   , 32  , 32
        BitsPerBlock = EConfigBitPerBlockUsed   ''''512  , 512 , 512
    
    ElseIf (FuseType = "UID") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = UIDBitsPerRow              ''''32   , 16  , 32
        ReadCycles = UIDReadCycle               ''''64   , 64  , 32
        BitsPerCycle = UIDBitsPerCycle          ''''32   , 32  , 32
        BitsPerBlock = UIDBitsPerBlockUsed      ''''1024 , 1024, 1024

    ElseIf (FuseType = "SEN") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = SENSORBitsPerRow           ''''32   , 16  , 32
        ReadCycles = SENSORReadCycle            ''''32   , 32  , 32
        BitsPerCycle = SENSORReadBitWidth       ''''32   , 32  , 32
        BitsPerBlock = SENSORBitPerBlockUsed    ''''512  , 512 , 1024

    ElseIf (FuseType = "MON") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = MONITORBitsPerRow           ''''32   , 16  , 32
        ReadCycles = MONITORReadCycle            ''''32   , 32  , 32
        BitsPerCycle = MONITORReadBitWidth       ''''32   , 32  , 32
        BitsPerBlock = MONITORBitPerBlockUsed    ''''512  , 512 , 1024

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,SEN,MON)"
        GoTo errHandler
        ''''nothing
    End If
    
    ''''20161114 update <MUST>
    ReDim SingleBitArray(ReadCycles * BitsPerCycle - 1)
    
    ''''    SingleStrArray() is [MSB.....LSB]
    '''' => Reverse SingleStrArray() to get SingleBitArray()
    '''' => SingleBitArray() is [LSB=bit0, ....., MSB=lastbit]
    
    singleBitSum = 0 ''''MUST
    
    If ((gS_EFuse_Orientation = "UP2DOWN") Or (gS_EFuse_Orientation = "SingleUp")) Then
        
        cnt = 0
        For i = 0 To ReadCycles - 1
            For j = 0 To BitsPerCycle - 1 ''''was BitsPerRow
                SingleBitArray(cnt) = CLng(Mid(StrReverse(SingleStrArray(i, ss)), j + 1, 1)) ''''StrReverse
                singleBitSum = singleBitSum + SingleBitArray(cnt)
                cnt = cnt + 1
            Next j
        Next i
        
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        ''''20170217 update
        For i = 0 To ReadCycles - 1
            For j = 0 To BitsPerCycle - 1
                tmpStr = Mid(StrReverse(SingleStrArray(i, ss)), j + 1, 1)
                SingleBitArray(i * BitsPerCycle + j) = CLng(IIf(tmpStr = "", "0", tmpStr))
                singleBitSum = singleBitSum + SingleBitArray(i * BitsPerCycle + j)
            Next j
        Next i
''''        If LCase(TheExec.DataManager.InstanceName) Like "*jtag*" Then
''''            For i = 0 To ReadCycles / 2 - 1
''''                For j = 0 To BitsPerCycle - 1
''''                    TmpStr = Mid(StrReverse(SingleStrArray(i, ss)), j + 1, 1)
''''                    SingleBitArray(i * BitsPerCycle + j) = CLng(IIf(TmpStr = "", "0", TmpStr))
''''                    SingleBitSum = SingleBitSum + SingleBitArray(i * BitsPerCycle + j)
''''                Next j
''''            Next i
''''        Else
''''            For i = 0 To ReadCycles - 1
''''                For j = 0 To BitsPerCycle - 1
''''                    TmpStr = Mid(StrReverse(SingleStrArray(i, ss)), j + 1, 1)
''''                    SingleBitArray(i * BitsPerCycle + j) = CLng(IIf(TmpStr = "", "0", TmpStr))
''''                    SingleBitSum = SingleBitSum + SingleBitArray(i * BitsPerCycle + j)
''''                Next j
''''            Next i
''''        End If
    End If
  
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_Gen_DoubleBitArray(ByVal FuseType As String, SingleBitArray() As Long, DoubleBitArray() As Long, doubleBitSum As Long)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Gen_DoubleBitArray"
    
    Dim i As Long, j As Long, k As Long
    Dim k1 As Long, k2 As Long
    Dim kk As Long
    
    Dim BitsPerRow As Long
    Dim ReadCycles As Long
    Dim BitsPerCycle As Long
    Dim BitsPerBlock As Long
    
    FuseType = UCase(Trim(FuseType))
    If (FuseType = "ECID") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = EcidBitsPerRow             ''''32   , 16  , 32
        ReadCycles = EcidReadCycle              ''''16   , 16  , 16
        BitsPerCycle = ECIDBitPerCycle          ''''32   , 32  , 32
        BitsPerBlock = EcidBitPerBlockUsed      ''''256  , 256 , 512
        
    ElseIf (FuseType = "CFG") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = EConfigBitsPerRow          ''''32   , 16  , 32
        ReadCycles = EConfigReadCycle           ''''32   , 32  , 16
        BitsPerCycle = EConfigReadBitWidth      ''''32   , 32  , 32
        BitsPerBlock = EConfigBitPerBlockUsed   ''''512  , 512 , 512
    
    ElseIf (FuseType = "UID") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = UIDBitsPerRow              ''''32   , 16  , 32
        ReadCycles = UIDReadCycle               ''''64   , 64  , 32
        BitsPerCycle = UIDBitsPerCycle          ''''32   , 32  , 32
        BitsPerBlock = UIDBitsPerBlockUsed      ''''1024 , 1024, 1024

    ElseIf (FuseType = "SEN") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = SENSORBitsPerRow           ''''32   , 16  , 32
        ReadCycles = SENSORReadCycle            ''''32   , 32  , 16
        BitsPerCycle = SENSORReadBitWidth       ''''32   , 32  , 32
        BitsPerBlock = SENSORBitPerBlockUsed    ''''512  , 512 , 512

    ElseIf (FuseType = "MON") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = MONITORBitsPerRow           ''''32   , 16  , 32
        ReadCycles = MONITORReadCycle            ''''32   , 32  , 16
        BitsPerCycle = MONITORReadBitWidth       ''''32   , 32  , 32
        BitsPerBlock = MONITORBitPerBlockUsed    ''''512  , 512 , 512

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,SEN,MON)"
        GoTo errHandler
        ''''nothing
    End If

    ''''<Important>
    ReDim DoubleBitArray(BitsPerBlock - 1)

    doubleBitSum = 0 ''''MUST
    
    If (gS_EFuse_Orientation = "UP2DOWN") Then

        For k = 0 To BitsPerBlock - 1 ''0...255(ECID), 0...511(CFG,SEN)
            DoubleBitArray(k) = SingleBitArray(k) Or SingleBitArray(k + BitsPerBlock)
            doubleBitSum = doubleBitSum + DoubleBitArray(k)
        Next k

    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
    
        k = 0 ''''must be here
        For i = 0 To ReadCycles - 1      ''0...15(ECID), 0...31(CFG,SEN)
            For j = 0 To BitsPerRow - 1  ''0...15(ECID), 0...15(CFG,SEN)
                ''''k1: Right block
                ''''k2:  Left block
                k1 = (i * BitsPerCycle) + j ''<Important> Must use BitsPerCycle here
                k2 = (i * BitsPerCycle) + BitsPerRow + j
                DoubleBitArray(k) = SingleBitArray(k1) Or SingleBitArray(k2)
                doubleBitSum = doubleBitSum + DoubleBitArray(k)
                k = k + 1
            Next j
        Next i
    
    ElseIf (gS_EFuse_Orientation = "SingleUp") Then

        ''''DoubleBitArray is equal to SingleBitArray
        For k = 0 To BitsPerBlock - 1 ''0...511(ECID)
            DoubleBitArray(k) = SingleBitArray(k)
            doubleBitSum = doubleBitSum + DoubleBitArray(k)
        Next k

    ''''Below is reserved for the future, there is NO any definition at present.
    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
    End If
   
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''It's for all bits
Public Function auto_EcidCompare_DoubleBit_PgmBit(DoubleBitArray() As Long, eFuse_Pgm_Bit() As Long, FailCnt As SiteLong)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_EcidCompare_DoubleBit_PgmBit"
    
    Dim ss As Variant
    Dim k As Long, j As Long
    
    ss = TheExec.sites.SiteNumber
    FailCnt(ss) = 0
    
    If (gS_EFuse_Orientation = "SingleUp") Then
        For k = 0 To EcidBitPerBlock - 1
            If DoubleBitArray(k) <> eFuse_Pgm_Bit(k) Then
                FailCnt(ss) = FailCnt(ss) + 1
            End If
        Next k
    ElseIf (gS_EFuse_Orientation = "UP2DOWN") Then
        ''''Up-Side
        For k = 0 To EcidBitPerBlock - 1
            If DoubleBitArray(k) <> eFuse_Pgm_Bit(k) Then
                FailCnt(ss) = FailCnt(ss) + 1
            End If
        Next k
        ''''Down-Side
        For k = 0 To EcidBitPerBlock - 1
            If DoubleBitArray(k) <> eFuse_Pgm_Bit(k + EcidBitPerBlockUsed) Then
                FailCnt(ss) = FailCnt(ss) + 1
            End If
        Next k
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        For k = 0 To EcidRowPerBlock - 1    ''0~15
            For j = 0 To EcidBitsPerRow - 1 ''0~15, 16 bits per row
                ''''Right-Side
                If DoubleBitArray(k * EcidBitsPerRow + j) <> eFuse_Pgm_Bit(k * EcidReadBitWidth + j) Then
                    FailCnt(ss) = FailCnt(ss) + 1
                End If
                ''''Left-Side
                If DoubleBitArray(k * EcidBitsPerRow + j) <> eFuse_Pgm_Bit(k * EcidReadBitWidth + EcidBitsPerRow + j) Then
                    FailCnt(ss) = FailCnt(ss) + 1
                End If
            Next j
        Next k
    End If
   
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150603, currently it's NOT used.
''''By this way, it can be easy to compare DoubleBit and PgmBit in the specific programming stage
Public Function auto_EcidCompare_DoubleBit_PgmBit_byStage(DoubleBitArray() As Long, eFuse_Pgm_Bit() As Long, FailCnt As SiteLong)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_EcidCompare_DoubleBit_PgmBit_byStage"

    Dim ss As Variant
    Dim i As Long, k As Long, j As Long
    Dim m_stage As String
    Dim m_startbit As Long
    Dim m_endbit As Long
    Dim m_bitwidth As Long
    Dim bcnt As Long

    ss = TheExec.sites.SiteNumber
    FailCnt(ss) = 0

    For i = 0 To UBound(ECIDFuse.Category) - 1 ''''skip the last one "DEID"
        m_stage = LCase(ECIDFuse.Category(i).Stage) ''''<Notice>

        If (gS_JobName = m_stage) Then
            If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
                m_startbit = ECIDFuse.Category(i).MSBbit
                m_endbit = ECIDFuse.Category(i).LSBbit
            Else
                ''''20160115, S4E startbit=LSBbit
                m_startbit = ECIDFuse.Category(i).LSBbit
                m_endbit = ECIDFuse.Category(i).MSBbit
            End If
            m_bitwidth = ECIDFuse.Category(i).BitWidth

            ''''------------------------------------------------------------
            If (gS_EFuse_Orientation = "SingleUp") Then
                For k = m_startbit To m_endbit
                    If DoubleBitArray(k) <> eFuse_Pgm_Bit(k) Then
                        FailCnt(ss) = FailCnt(ss) + 1
                    End If
                Next k
            ElseIf (gS_EFuse_Orientation = "UP2DOWN") Then
                For k = m_startbit To m_endbit
                    ''''Up-Side
                    If (DoubleBitArray(k) <> eFuse_Pgm_Bit(k)) Then
                        FailCnt(ss) = FailCnt(ss) + 1
                    End If
                    ''''Down-Side
                    If (DoubleBitArray(k) <> eFuse_Pgm_Bit(k + EcidBitPerBlockUsed)) Then
                        FailCnt(ss) = FailCnt(ss) + 1
                    End If
                Next k
            ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
                bcnt = 0 ''''<MUST>
                For k = 0 To EcidRowPerBlock - 1    ''0...15
                    For j = 0 To EcidBitsPerRow - 1 ''0...15, 16 bits per row
                        If (bcnt >= m_startbit And bcnt <= m_endbit) Then
                            ''''Right-Side
                            If DoubleBitArray(k * EcidBitsPerRow + j) <> eFuse_Pgm_Bit(k * EcidReadBitWidth + j) Then
                                FailCnt(ss) = FailCnt(ss) + 1
                            End If
                            ''''Left-Side
                            If DoubleBitArray(k * EcidBitsPerRow + j) <> eFuse_Pgm_Bit(k * EcidReadBitWidth + EcidBitsPerRow + j) Then
                                FailCnt(ss) = FailCnt(ss) + 1
                            End If
                        Else
                            ''''over m_endbit bits, set k,j to up limit to escape for-loop
                            If (bcnt > m_endbit) Then
                                k = EcidRowPerBlock
                                j = EcidBitsPerRow
                            End If
                        End If
                        bcnt = bcnt + 1
                    Next j
                Next k
            End If
            ''''------------------------------------------------------------
        End If
    Next i

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160201, especially ECID Blank check DEID for the non-CP1 stage
Public Function auto_EcidPgmBit_DEID_forCheck(ByRef Expand_eFuse_Pgm_Bit() As Long, ByRef eFuse_Pgm_Bit() As Long, Optional showPrint As Boolean = True) As Long
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_EcidPgmBit_DEID_forCheck"
                    
    Dim i As Long, j As Long, k As Long
    Dim k1 As Long, k2 As Long
    Dim CountN As Long
    Dim LotBinaryStr As String
    Dim LotBinaryStr_R As String ''''StrReverse
    
    Dim kk As Long
    Dim idx As Long
    Dim tmpStr As String
    Dim m_stage As String
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_MSBBit As Long
    Dim m_LSBbit As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant ''was Long, 20170911 update
    Dim m_pgm_flag As Long
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim tmpdlgStr As String
    Dim m_bitstrL As String
    Dim m_bitStrM As String
    Dim m_tmpVal As Variant
    Dim m_HexStr As String

    Dim ss As Variant
    ss = TheExec.sites.SiteNumber
    
    '<<< Step1. make eFuse_Pgm_Bit >>>
    'Beware, due to single bit, double bit will consume two times for total bits
    'if programed efuse_pgm_bit over bit255, those bits are unprogramming.
    'need to check if last program bit still under total bits/2
    
    ''''Initialize
    For i = 0 To UBound(eFuse_Pgm_Bit)
        eFuse_Pgm_Bit(i) = 0
    Next i

    LotBinaryStr = ""
    idx = ECIDIndex("Lot_ID")
    m_MSBBit = ECIDFuse.Category(idx).MSBbit
    m_LSBbit = ECIDFuse.Category(idx).LSBbit
    m_stage = LCase(ECIDFuse.Category(idx).Stage)
    m_bitwidth = ECIDFuse.Category(idx).BitWidth
    m_defval = ECIDFuse.Category(idx).DefaultValue
    m_defreal = ECIDFuse.Category(idx).Default_Real

    ''''-----------------------------------------------------------------------------------------------
    ''''simulation for non-CP stage
    ''''-----------------------------------------------------------------------------------------------
    If ((TheExec.TesterMode = testModeOffline) And (gB_ReadWaferData_flag = False)) Then
        Dim m_tmpID As String
        Dim m_tmpwfid As String
        m_tmpID = TheExec.Datalog.Setup.LotSetup.LotID
        If (Len(m_tmpID) < 6) Then
            LotID = "DUMMY0"
            TheExec.Datalog.WriteComment vbTab & "<Offline> Set LotID = DUMMY0  (pseudo lotid)"
        Else
            LotID = Mid(TheExec.Datalog.Setup.LotSetup.LotID, 1, 6)
        End If
        m_tmpwfid = TheExec.Datalog.Setup.WaferSetup.ID
        If (IsNumeric(CStr(m_tmpwfid)) = False) Then
            WaferID = 25
            TheExec.Datalog.WriteComment vbTab & "<Offline> Set WaferID = 25 (pseudo wafer id)"
        Else
            WaferID = CLng(m_tmpwfid)
        End If
        ''XCoord(ss) = theExec.Datalog.Setup.WaferSetup.GetXCoord(ss)
        ''YCoord(ss) = theExec.Datalog.Setup.WaferSetup.GetYCoord(ss)
        ''If (XCoord(ss) = -32768 Or YCoord(ss) = -32768) Then
        If (True) Then
            ''Call setXY(5, 6) ''''set a pseudo XY coordinate
            ''theExec.Datalog.WriteComment vbTab & "Call setXY(5, 6) (pseudo XY_Coordinate)"
            ''XCoord(ss) = theExec.Datalog.Setup.WaferSetup.GetXCoord(ss)
            ''YCoord(ss) = theExec.Datalog.Setup.WaferSetup.GetYCoord(ss)

            ''''20160525 update
            XCoord(ss) = 1 + Int(1 + Rnd(1) * 11) + ss
            YCoord(ss) = 2 + Int(2 + Rnd(2) * 12) + ss
            
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(ss, XCoord(ss))
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(ss, YCoord(ss))
            TheExec.Datalog.WriteComment vbTab & "<Offline> Set XCoord(" + FormatNumeric(ss, 1) + ") = " + FormatNumeric(XCoord(ss), -3) + " (pseudo XCoord)"
            TheExec.Datalog.WriteComment vbTab & "<Offline> Set YCoord(" + FormatNumeric(ss, 1) + ") = " + FormatNumeric(YCoord(ss), -3) + " (pseudo YCoord)"
        End If
        Call auto_eFuse_SetWriteDecimal("ECID", "Wafer_ID", WaferID, False, False)
        Call auto_eFuse_SetWriteDecimal("ECID", "X_Coordinate", XCoord(ss), False, False)
        Call auto_eFuse_SetWriteDecimal("ECID", "Y_Coordinate", YCoord(ss), False, False)
    End If
    ''''-----------------------------------------------------------------------------------------------
    
    ''If (gB_ReadWaferData_flag = True) Then ''''update
    ''If (gB_ReadWaferData_flag = True Or (TheExec.TesterMode = testModeOffline And gB_ReadWaferData_flag = False)) Then ''''update
    If (True) Then ''''20170829 update
        ''''MSB First = Y means that the 1st bit (bit0) is MSB
        If (UCase(ECIDFuse.Category(idx).MSBFirst) = "Y") Then
            ''''           char1          2          3          4          5          6
            ''''LotBinaryStr = [MSB...LSB][MSB...LSB][MSB...LSB][MSB...LSB][MSB...LSB][MSB...LSB]
            For i = 1 To EcidCharPerLotId
                LotBinaryStr = LotBinaryStr + auto_MappingCharToBinStr(Mid(LotID, i, 1))
            Next i
            ''''BitStrM:: BitString MSB......LSB
            ''''BitStrL:: BitString LSB......MSB
            ECIDFuse.Category(idx).Write.BitStrM(ss) = LotBinaryStr
            ECIDFuse.Category(idx).Write.BitStrL(ss) = StrReverse(LotBinaryStr)
            ECIDFuse.Category(idx).Write.Value(ss) = LotID
            ECIDFuse.Category(idx).Write.ValStr(ss) = LotID

            k = 0  ''k=0 stand for bit 31 of first words
            For i = 0 To EcidCharPerLotId - 1   'LotId 6 char
                For j = 0 To EcidBitPerLotIdChar - 1 '6bit per char
                    k = m_MSBBit + (i * EcidBitPerLotIdChar + j)
                    kk = i * EcidBitPerLotIdChar + j
                    eFuse_Pgm_Bit(k) = CLng(Mid(LotBinaryStr, kk + 1, 1)) ''EcidCharKey("0")="000000" LotIdCharMapping
                    'Debug.Print i & "," & j & ": eFuse_Pgm_Bit(" & k & ") = " & eFuse_Pgm_Bit(k)
                Next j
            Next i
        Else
            ''''MSBFirst='N' or empty
            ''''EX: LotID='NP5678'
            ''''           char1(N)       2(P)       3(5)       4(6)       5(7)       6(8)
            ''''LotBinaryStr = [MSB...LSB][MSB...LSB][MSB...LSB][MSB...LSB][MSB...LSB][MSB...LSB]
            For i = 1 To EcidCharPerLotId
                LotBinaryStr = LotBinaryStr + auto_MappingCharToBinStr(Mid(LotID, i, 1))
            Next i
            ''''BitStrM:: BitString MSB......LSB
            ''''BitStrL:: BitString LSB......MSB
            ECIDFuse.Category(idx).Write.BitStrM(ss) = LotBinaryStr
            ECIDFuse.Category(idx).Write.BitStrL(ss) = StrReverse(LotBinaryStr)
            ECIDFuse.Category(idx).Write.Value(ss) = LotID
            ECIDFuse.Category(idx).Write.ValStr(ss) = LotID

            LotBinaryStr_R = StrReverse(LotBinaryStr) ''''to [LSB......MSB]
            kk = 0
            k = 0
            For i = 0 To EcidCharPerLotId - 1   'LotId 6 char
                For j = 0 To EcidBitPerLotIdChar - 1 '6bit per char
                    k = m_LSBbit + i * EcidBitPerLotIdChar + j
                    kk = i * EcidBitPerLotIdChar + j
                    eFuse_Pgm_Bit(k) = CLng(Mid(LotBinaryStr_R, kk + 1, 1))
                    ''Debug.Print i & "," & j & ": eFuse_Pgm_Bit(" & k & ") = " & eFuse_Pgm_Bit(k)
                Next j
            Next i
        End If
    Else
        For j = 1 To m_bitwidth
            m_bitstrL = m_bitstrL + "0"
        Next j
        ECIDFuse.Category(idx).Write.BitStrM(ss) = m_bitstrL ''''all zero
        ECIDFuse.Category(idx).Write.BitStrL(ss) = m_bitstrL
        ECIDFuse.Category(idx).Write.Value(ss) = "000000"
        ECIDFuse.Category(idx).Write.ValStr(ss) = "000000"

    End If ''''end of If (m_stage = gS_JobName And gB_ReadWaferData_flag = True)

    ''''-----------------------------------------------------------------------------------------------------
    ''''Only Programming first DEID,DAY,DEVICE,FUSE and reserved bits
    ''''-----------------------------------------------------------------------------------------------------
    ''''<NOTICE> Here this function only check DEID only (LotID/WaferID/X/Y)
    If (showPrint) Then TheExec.Datalog.WriteComment ""
    For i = 0 To UBound(ECIDFuse.Category) - 1 ''''skip the last one algorithm "DEID"
        tmpStr = ""
        m_pgm_flag = 0 ''''<MUST>
        m_catename = ECIDFuse.Category(i).Name
        m_stage = LCase(ECIDFuse.Category(i).Stage)
        m_algorithm = LCase(ECIDFuse.Category(i).algorithm)
        m_MSBBit = ECIDFuse.Category(i).MSBbit
        m_LSBbit = ECIDFuse.Category(i).LSBbit
        m_bitwidth = ECIDFuse.Category(i).BitWidth
        m_defval = ECIDFuse.Category(i).DefaultValue
        m_defreal = ECIDFuse.Category(i).Default_Real

        ''''20150710 new datalog format
        tmpdlgStr = "Site(" + CStr(ss) + ") Simulate : " + FormatNumeric(m_catename, gI_ECID_catename_maxLen)
        If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
            If (m_algorithm = "crc") Then
                ''''<NOTICE>20170331, ECID CRC always coding/fusing as [MSB......LSB]
                ''''Example: m_LSBbit(255) is MSB of CRC result, m_MSBbit(240) is LSB of CRC result
                ''''         gL_ECID_CRC_MSB = m_LSBbit(255), gL_ECID_CRC_LSB = m_MSBbit(240)
                ''''The above is correct in the function auto_ECIDConstant_Initialize()
                tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(gL_ECID_CRC_MSB, "0000") + ":" + Format(gL_ECID_CRC_LSB, "0000") + "(LSB)] = "
            Else
                tmpdlgStr = tmpdlgStr + " [(LSB)" + Format(m_LSBbit, "0000") + ":" + Format(m_MSBBit, "0000") + "(MSB)] = "
            End If
        Else
            tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "(LSB)] = "
        End If

        m_decimal = 0 ''''Must be here for the initilization every time
        If (m_algorithm = "lotid") Then
            m_pgm_flag = 1 ''''<NOTICE> MUST be 1, and it has been programmed above.
        ElseIf (m_algorithm = "numeric") Then
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("ECID", m_catename, m_defreal, m_defval, False)
            m_pgm_flag = 1
        ElseIf (m_algorithm Like "*reserve*") Then
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("ECID", m_catename, m_defreal, m_defval, False)
            m_pgm_flag = 1
        ElseIf (m_algorithm = "crc") Then
            m_pgm_flag = 0 ''''update below, 20170829

        Else
            ''''20160201 New, for the non-CP1 simulation of S4E
            If ((TheExec.TesterMode = testModeOffline) And (gS_JobName <> "cp1")) Then
                m_pgm_flag = 1
            End If
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("ECID", m_catename, m_defreal, m_defval, False)
        End If

        ''''-----------------------------------------------------------
        ''''new, 20170911 update
        ''''-----------------------------------------------------------
        If ((m_pgm_flag = 1) Or (gB_ReadWaferData_flag = True)) Then
            ''''<NOTICE> because lotid has been programmed above.
            ''''                 crc will be programmed late as below
            If (m_algorithm <> "lotid" And m_algorithm <> "crc") Then
                ''''20170911 update
                If (auto_isHexString(CStr(m_decimal))) Then
                    If (auto_chkHexStr_isOver7FFFFFFF(CStr(m_decimal)) = True) Then
                        Call auto_eFuse_HexStr2PgmArr_Write_byStage("ECID", m_stage, CLng(i), m_decimal, m_LSBbit, m_MSBBit, eFuse_Pgm_Bit, False) ''''<MUST>set chkStage=False
                    Else
                        ''''20170911 update
                        If (UCase(CStr(m_decimal)) Like "0X*") Then
                            m_decimal = Replace(UCase(CStr(m_decimal)), "0X", "", 1, 1)
                        ElseIf (UCase(CStr(m_decimal)) Like "X*") Then
                            m_decimal = Replace(UCase(CStr(m_decimal)), "X", "", 1, 1)
                        End If
                        m_decimal = CLng("&H" & CStr(m_decimal)) ''''Here it's Hex2Dec
                        If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
                            Call auto_CalcEfuseBit_Reversed(i, CLng(m_decimal), m_MSBBit, m_bitwidth, eFuse_Pgm_Bit, tmpStr)
                        Else
                            Call auto_CalcEfuseBit(i, CLng(m_decimal), m_LSBbit, m_bitwidth, eFuse_Pgm_Bit, tmpStr)
                        End If
                    End If
                Else
                    ''''not a hex String
                    If (m_decimal <= (CDbl(2 ^ 31) - 1)) Then ''''<= 0x7FFFFFFF
                        If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
                            Call auto_CalcEfuseBit_Reversed(i, CLng(m_decimal), m_MSBBit, m_bitwidth, eFuse_Pgm_Bit, tmpStr)
                        Else
                            Call auto_CalcEfuseBit(i, CLng(m_decimal), m_LSBbit, m_bitwidth, eFuse_Pgm_Bit, tmpStr)
                        End If
                    Else
                        ''''over Long range (> 0x7FFFFFFF)
                        ''''Firstly, convert to Hex String with prefix '0x'
                        m_HexStr = auto_Value2HexStr(m_decimal, m_bitwidth)
                        Call auto_eFuse_HexStr2PgmArr_Write_byStage("ECID", m_stage, CLng(i), m_HexStr, m_LSBbit, m_MSBBit, eFuse_Pgm_Bit, False) ''''<MUST>set chkStage=False
                    End If
                End If
            End If
        Else
            ''''<Notes> must skip "DEID" to avoid all first DEID are cleaned.
            If (m_algorithm <> LCase("DEID") And m_algorithm <> "crc") Then ''''20170829 skip "crc"
                m_decimal = 0 ''''<MUST>
                If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
                    Call auto_CalcEfuseBit_Reversed(i, CLng(m_decimal), m_MSBBit, m_bitwidth, eFuse_Pgm_Bit, tmpStr)
                Else
                    Call auto_CalcEfuseBit(i, CLng(m_decimal), m_LSBbit, m_bitwidth, eFuse_Pgm_Bit, tmpStr)
                End If
            End If
        End If
        ''''-----------------------------------------------------------

        ''''20150710 new
        If (m_algorithm <> LCase("DEID") And m_algorithm <> "crc") Then ''''skip "ECID_DEID", "crc", ''''20170829 update
            m_tmpVal = ECIDFuse.Category(i).Write.Value(ss)
            m_bitstrL = ECIDFuse.Category(i).Write.BitStrL(ss)
            m_bitStrM = ECIDFuse.Category(i).Write.BitStrM(ss)
            m_HexStr = ECIDFuse.Category(i).Write.HexStr(ss)
            If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
                tmpStr = " [" + m_bitstrL + "]"
            Else
                tmpStr = " [" + m_bitStrM + "]"
            End If
            
            ''''20170911 update
            If (m_bitwidth < 16 Or m_algorithm = "lotid") Then
                tmpdlgStr = tmpdlgStr + FormatNumeric(m_tmpVal, 10) + tmpStr
            Else
                tmpdlgStr = tmpdlgStr + FormatNumeric(m_tmpVal, 10) + tmpStr + FormatNumeric(" [" + m_HexStr + "]", -1)
            End If
            If (showPrint) Then TheExec.Datalog.WriteComment tmpdlgStr
        End If
    Next i
    
    ''''20170829 update
    ''''calc CRC from the eFuse_Pgm_Bit with gL_ECID_CRC_calcBits()
    For i = 0 To UBound(ECIDFuse.Category) - 1 ''''skip the last one algorithm "DEID"
        tmpStr = ""
        m_pgm_flag = 0 ''''<MUST>
        m_catename = ECIDFuse.Category(i).Name
        m_stage = LCase(ECIDFuse.Category(i).Stage)
        m_algorithm = LCase(ECIDFuse.Category(i).algorithm)
        m_MSBBit = ECIDFuse.Category(i).MSBbit
        m_LSBbit = ECIDFuse.Category(i).LSBbit
        m_bitwidth = ECIDFuse.Category(i).BitWidth
        m_defval = ECIDFuse.Category(i).DefaultValue
        m_defreal = ECIDFuse.Category(i).Default_Real
    
        If (m_algorithm = "crc") Then
            ''''20150710 new datalog format
            tmpdlgStr = "Site(" + CStr(ss) + ") Simulate : " + FormatNumeric(m_catename, gI_ECID_catename_maxLen)
            If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
                If (m_algorithm = "crc") Then
                    ''''<NOTICE>20170331, ECID CRC always coding/fusing as [MSB......LSB]
                    ''''Example: m_LSBbit(255) is MSB of CRC result, m_MSBbit(240) is LSB of CRC result
                    ''''         gL_ECID_CRC_MSB = m_LSBbit(255), gL_ECID_CRC_LSB = m_MSBbit(240)
                    ''''The above is correct in the function auto_ECIDConstant_Initialize()
                    tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(gL_ECID_CRC_MSB, "0000") + ":" + Format(gL_ECID_CRC_LSB, "0000") + "(LSB)] = "
                Else
                    tmpdlgStr = tmpdlgStr + " [(LSB)" + Format(m_LSBbit, "0000") + ":" + Format(m_MSBBit, "0000") + "(MSB)] = "
                End If
            Else
                tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "(LSB)] = "
            End If

            Dim CRCarray(15) As Byte
            Dim crcStr As String
            Dim CRCHex As String
            crcStr = ""
            Call CRC_Zero_Array(CRCarray)
            gL_CRCidx = 0
            For j = UBound(gL_ECID_CRC_calcBits) To 0 Step -1 ''''bit63 to bit0
                If (gL_ECID_CRC_calcBits(j) = 1) Then
                    Call CRC16_ComputeCRCforBit(CRCarray, CByte(eFuse_Pgm_Bit(j)), False) ''''set True for the debug.
                End If
            Next j

            ''''20161004
            ''''<NOTICE> CRCarray(0) should be LSB bit
            crcStr = ""
            For j = 0 To UBound(CRCarray)
                eFuse_Pgm_Bit(gL_ECID_CRC_LSB + j) = CRCarray(j) ''''Tmys gL_ECID_CRC_LSB=ECIDFuse.Category(i).MSBbit
                ''''eFuse_Pgm_Bit(gL_ECID_CRC_LSB - j) = CRCarray(j)
                ''crcStr = crcStr & CRCarray(j)
                crcStr = CRCarray(j) & crcStr   ''''[MSB......LSB], 20170331 update
            Next j

            '''''''''''1111110000000000
            '''''''''''5432109876543210
            'CRCStr = "0000000000000000"
            CRCHex = auto_BinStr2HexStr(crcStr, 4)
            ECIDFuse.Category(i).Write.BitStrM(ss) = crcStr
            ECIDFuse.Category(i).Write.BitStrL(ss) = StrReverse(crcStr)
            ECIDFuse.Category(i).Write.Value(ss) = CRCHex
            ECIDFuse.Category(i).Write.ValStr(ss) = CRCHex
            ECIDFuse.Category(i).Write.HexStr(ss) = "0x" + CStr(CRCHex)
            m_pgm_flag = 1

            m_tmpVal = ECIDFuse.Category(i).Write.HexStr(ss)        ''''20170331 update for CRC
            tmpStr = " [" + crcStr + "]"                            ''''CRC always [MSB....LSB]
            tmpdlgStr = tmpdlgStr + FormatNumeric(m_tmpVal, 10) + tmpStr
            If (showPrint) Then TheExec.Datalog.WriteComment tmpdlgStr
            Exit For ''''<MUST> to save time
        End If
    Next i

    If (showPrint) Then TheExec.Datalog.WriteComment ""
    ''''-----------------------------------------------------------------------------------------------------

    '<<< Step1.5 reorder eFuse_Pgm_Bit, Up to Down>>>
    If (gS_EFuse_Orientation = "UP2DOWN") Then
        'copy from eFuse_Pgm_Bit(Up) to eFuse_Pgm_Bit(Down)
        For i = 0 To EcidBitPerBlockUsed - 1
            eFuse_Pgm_Bit(i + EcidBitPerBlockUsed) = eFuse_Pgm_Bit(i)
        Next i
    End If
    
    '<<< Step1.5 reorder eFuse_Pgm_Bit, Right to Left>>>
    If (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        Dim TempeFuseArray() As Long
        ReDim TempeFuseArray(EcidBitPerBlockUsed)
        'copy from eFuse_Pgm_Bit
        For i = 0 To EcidBitPerBlockUsed - 1
            TempeFuseArray(i) = eFuse_Pgm_Bit(i)
        Next i
        
        ''''-------------------------------------------------------------------------------
        ''''New Method
        ''''-------------------------------------------------------------------------------
        k = 0 ''''must be here
        For i = 0 To EcidReadCycle - 1       ''0...15(ECID)
            For j = 0 To EcidBitsPerRow - 1  ''0...15(ECID)
                ''''k1: Right block
                ''''k2:  Left block
                k1 = (i * EcidReadBitWidth) + j ''<Important> Must use EcidReadBitWidth here
                k2 = (i * EcidReadBitWidth) + EcidBitsPerRow + j
                eFuse_Pgm_Bit(k1) = TempeFuseArray(k)
                eFuse_Pgm_Bit(k2) = TempeFuseArray(k)
                k = k + 1
            Next j
        Next i
        ''''-------------------------------------------------------------------------------
    End If

    'Multiple EcidWriteBitExpandWidth times for DSSC wave
    CountN = 0
    For i = 0 To UBound(eFuse_Pgm_Bit)
        For j = 0 To EcidWriteBitExpandWidth - 1
            Expand_eFuse_Pgm_Bit(CountN) = eFuse_Pgm_Bit(i)
            CountN = CountN + 1
        Next j
    Next i

    auto_EcidPgmBit_DEID_forCheck = CountN

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

''''For the case:: Ecid MSB_First is 'Y'
''''Mean that ECID_Array 1st element is MSB
''Was:: Public Function auto_EcidConvBit2NumStr(ByRef field As SiteLong, num_bit As Long, ByRef ECID_Array() As String, ByRef start_bit As Long) 'site dependent
Public Function auto_EcidConvBit2NumStr(idx As Long, start_bit As Long, num_bit As Long, ByRef ECID_Array() As String, ByRef decValue As Long) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_EcidConvBit2NumStr"
    
    Dim m_decimal As Long
    Dim Cnt1 As Long
    Dim i As Long
    Dim m_bitStrM As String
    Dim m_bitsum As Long
    
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber
    
    m_bitStrM = ""
    m_decimal = 0
    Cnt1 = start_bit
    m_bitsum = 0
    
    ''''if num_bit >=31, it will result in an overflow error message and supposedy it's the reserved bits.
    If (num_bit <= 31) Then
        For i = 0 To num_bit - 1
            m_bitStrM = m_bitStrM + CStr(ECID_Array(Cnt1)) ''''[MSB......LSB]
            If CInt(ECID_Array(Cnt1)) = 1 Then
                m_decimal = m_decimal + 2 ^ (num_bit - 1 - i)
            End If
            m_bitsum = m_bitsum + CInt(ECID_Array(Cnt1))
            Cnt1 = Cnt1 + 1
        Next i
    Else
        ''''supposedy it's the reserved bits.
        m_decimal = 0
        For i = 0 To num_bit - 1
            m_bitStrM = m_bitStrM + CStr(ECID_Array(Cnt1)) ''''[MSB......LSB]
            m_bitsum = m_bitsum + CInt(ECID_Array(Cnt1))
        Next i
    End If
    
    ''''new, 20150529
    ECIDFuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
    ECIDFuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
    ECIDFuse.Category(idx).Read.Decimal(ss) = m_decimal
    ECIDFuse.Category(idx).Read.Value(ss) = m_decimal
    ECIDFuse.Category(idx).Read.ValStr(ss) = CStr(m_decimal)
    ECIDFuse.Category(idx).Read.BitSummation(ss) = m_bitsum

    decValue = m_decimal
    auto_EcidConvBit2NumStr = m_bitStrM
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''For the case:: Ecid MSB_First is 'N' or '' (empty)
''''Mean that ECID_Array 1st element is LSB
Public Function auto_EcidConvBit2NumStr_L(idx As Long, start_bit As Long, num_bit As Long, ByRef ECID_Array() As String, ByRef decValue As Long) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_EcidConvBit2NumStr_L"
    
    Dim m_decimal As Long
    Dim Cnt1 As Long
    Dim i As Long
    Dim m_bitStrM As String
    Dim m_bitsum As Long

    Dim ss As Variant
    ss = TheExec.sites.SiteNumber

    m_bitStrM = ""
    m_decimal = 0
    Cnt1 = start_bit
    
    ''''if num_bit >=31, it will result in an overflow error message and supposedy it's the reserved bits.
    If (num_bit <= 31) Then
        For i = 0 To num_bit - 1
            m_bitStrM = CStr(ECID_Array(Cnt1)) + m_bitStrM ''''[MSB......LSB]
            If CInt(ECID_Array(Cnt1)) = 1 Then
                m_decimal = m_decimal + 2 ^ (i)
            End If
            m_bitsum = m_bitsum + CInt(ECID_Array(Cnt1))
            Cnt1 = Cnt1 + 1
        Next i
    Else
        ''''supposedy it's the reserved bits.
        m_decimal = 0
        For i = 0 To num_bit - 1
            m_bitStrM = CStr(ECID_Array(Cnt1)) + m_bitStrM ''''[MSB......LSB]
            m_bitsum = m_bitsum + CInt(ECID_Array(Cnt1))
        Next i
    End If

    ''''new, 20150529
    ECIDFuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
    ECIDFuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
    ECIDFuse.Category(idx).Read.Decimal(ss) = m_decimal
    ECIDFuse.Category(idx).Read.Value(ss) = m_decimal
    ECIDFuse.Category(idx).Read.ValStr(ss) = CStr(m_decimal)
    ECIDFuse.Category(idx).Read.BitSummation(ss) = m_bitsum

    decValue = m_decimal
    auto_EcidConvBit2NumStr_L = m_bitStrM
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''nonDEID=non DeviceID
Public Function auto_BlankChk_nonDEID(SingleBitArray() As Long, ByRef blank As SiteBoolean, Optional SingleDoubleFBC As Long)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_BlankChk_nonDEID"
    
    Dim ss As Variant
    Dim k As Long, j As Long, i As Long
    Dim k1 As Long, k2 As Long
    Dim bcnt As Long
    Dim SingleSum As Long, DoubleSum As Long
    Dim TempDoubleBit As Long
    
    Dim m_match_flag As Boolean
    Dim m_algorithm As String
    Dim m_startbit As Long
    Dim m_endbit As Long
    
    SingleSum = 0: DoubleSum = 0
    
    ss = TheExec.sites.SiteNumber
    blank(ss) = True

    For i = 0 To UBound(ECIDFuse.Category) - 1 ''''<MUST> skip the last one "DEID"
        m_algorithm = LCase(ECIDFuse.Category(i).algorithm)
        If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
            m_startbit = ECIDFuse.Category(i).MSBbit
            m_endbit = ECIDFuse.Category(i).LSBbit
        Else
            m_startbit = ECIDFuse.Category(i).LSBbit
            m_endbit = ECIDFuse.Category(i).MSBbit
        End If
        m_match_flag = False ''''<MUST>

        If (TheExec.Flow.EnableWord("WAT_Enable") = True) Then
            If (m_algorithm <> "lotid" And m_algorithm <> "numeric" And m_algorithm <> "wat" And m_algorithm <> "device") Then ''''20151228
                m_match_flag = True
            End If
        Else
            If (m_algorithm <> "lotid" And m_algorithm <> "numeric" And m_algorithm <> "crc") Then ''''20161005
                m_match_flag = True
            End If
        End If

        ''''Matching then get the blank,SingleDoubleFBC
        If (m_match_flag = True) Then
            If (gS_EFuse_Orientation = "SingleUp") Then
                For k = m_startbit To m_endbit
                    SingleSum = SingleSum + SingleBitArray(k)
                    DoubleSum = DoubleSum + SingleBitArray(k)
                    If (SingleBitArray(k) <> 0) Then
                        blank(ss) = False
                    End If
                Next k
                SingleDoubleFBC = DoubleSum - SingleSum
            
            ElseIf (gS_EFuse_Orientation = "UP2DOWN") Then
                For k = m_startbit To m_endbit
                    SingleSum = SingleSum + (SingleBitArray(k) + SingleBitArray(EcidBitPerBlock + k))
                    TempDoubleBit = SingleBitArray(k) Or SingleBitArray(EcidBitPerBlock + k)
                    DoubleSum = DoubleSum + TempDoubleBit
                    If (SingleBitArray(k) <> 0 Or SingleBitArray(EcidBitPerBlock + k) <> 0) Then
                        blank(ss) = False
                        ''Exit For
                    End If
                Next k
                SingleDoubleFBC = DoubleSum * 2 - SingleSum

            ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
                bcnt = 0 ''''<MUST>
                ''''total 16x32=512
                For k = 0 To EcidReadCycle - 1      ''0 to 15
                    For j = 0 To EcidBitsPerRow - 1 ''0 to 15
                        ''''k1: Right block
                        ''''k2:  Left block
                        k1 = (k * EcidReadBitWidth) + j ''<Important> Must use EcidReadBitWidth here
                        k2 = (k * EcidReadBitWidth) + EcidBitsPerRow + j
                        
                        If (bcnt >= m_startbit And bcnt <= m_endbit) Then
                            SingleSum = SingleSum + (SingleBitArray(k1) + SingleBitArray(k2))
                            TempDoubleBit = SingleBitArray(k1) Or SingleBitArray(k2)
                            DoubleSum = DoubleSum + TempDoubleBit
                            If (SingleBitArray(k1) <> 0 Or SingleBitArray(k2) <> 0) Then
                                blank(ss) = False
                                'Exit For
                            End If
                        Else
                            ''''over m_endbit bits, set k,j to up limit to escape for-loop
                            If (bcnt > m_endbit) Then
                                k = EcidReadCycle
                                j = EcidBitsPerRow
                            End If
                        End If
                        bcnt = bcnt + 1
                    Next j
                Next k
                SingleDoubleFBC = DoubleSum * 2 - SingleSum
            End If
        End If
    Next i
   
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20180727 New
Public Function auto_eFuse_LotID_to_setWriteVariable(m_lotid As Variant) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_LotID_to_setWriteVariable"
    
    Dim i As Long
    Dim m_Site As Variant
    Dim BinStr As String
    Dim m_binarr() As Long
    Dim m_tmpSiteVar As New SiteVariant
    Dim m_tmpWave As New DSPWave
    Dim m_bitstrL As String
    
    m_tmpSiteVar = m_lotid
    
    '< Step1. transfer lotid to binary string>
    ''''BinStr="Ch1[MSB...LSB] Ch2[MSB...LSB] Ch3[MSB...LSB] Ch4[MSB...LSB] Ch5[MSB...LSB] Ch6[MSB...LSB]"
    BinStr = ""
    BinStr = BinStr + auto_LotIDCh2Binary(CStr(m_lotid))

    ''''Here m_binarr(0) is the MSB_bit of the 1st Character of Lotid
    ReDim m_binarr(Len(BinStr) - 1)
    m_bitstrL = StrReverse(BinStr)
    
    ''''<NOTICE> here m_binarr(0) is LSB bit
    For i = 1 To Len(BinStr)
        m_binarr(i - 1) = CLng(Mid(m_bitstrL, i, 1))
    Next i
    
    i = ECIDIndex("Lot_ID")

    m_tmpWave.CreateConstant 0, ECIDFuse.Category(i).BitWidth, DspLong
    For Each m_Site In TheExec.sites
        m_tmpWave.Data = m_binarr
    Next m_Site

    With ECIDFuse.Category(i).Write
        .Decimal = m_tmpSiteVar
        .Value = m_tmpSiteVar
        .HexStr = "0x" + auto_Binary2Hex(BinStr)
        .ValStr = m_tmpSiteVar
        .BitStrM = BinStr
        .BitStrL = StrReverse(BinStr)
        For Each m_Site In TheExec.sites
            .BitArrWave = m_tmpWave.Copy
            .BitSummation = .BitArrWave.CalcSum
        Next m_Site
    End With

    auto_eFuse_LotID_to_setWriteVariable = 1

''''Debug
If (False) Then
    For i = 0 To 35
    Debug.Print i & "=" & ECIDFuse.Category(0).Write.BitArrWave(0).Element(i)
    Next i
End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20180726 New for SiteAware
Public Function auto_WaferData_to_HexECID_SiteAware(Optional bitwith As Long = 64) As SiteVariant

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_WaferData_to_HexECID_SiteAware"
    
    Dim idx0 As Long
    Dim idx1 As Long
    Dim idx2 As Long
    Dim idx3 As Long
    Dim m_Site As Variant
    Dim m_BinStr As String
    Dim m_HexStr As New SiteVariant
    Dim m_hexlen As Long
    
    idx0 = ECIDIndex("Lot_ID")
    idx1 = ECIDIndex("Wafer_ID")
    idx2 = ECIDIndex("X_Coordinate")
    idx3 = ECIDIndex("Y_Coordinate")
    ''''m_BinStr="Ch1[MSB...LSB] Ch2[MSB...LSB] Ch3[MSB...LSB] Ch4[MSB...LSB] Ch5[MSB...LSB] Ch6[MSB...LSB]"
    ''''           m_BinStr  = [Lot_ID][Wafer_ID][X_Coordinate][Y_Coordinate]  (MSB......LSB)
    ''''StrReverse(m_BinStr) = [Y_Coordinate][X_Coordinate][Wafer_ID][Lot_ID]  (LSB......MSB)
    m_hexlen = IIf((bitwith Mod 4) = 0, bitwith \ 4, 1 + (bitwith \ 4))

    For Each m_Site In TheExec.sites
'        m_BinStr = ""
'        m_BinStr = m_BinStr + ECIDFuse.Category(idx0).Write.BitstrM
'        m_BinStr = m_BinStr + ECIDFuse.Category(idx1).Write.BitstrM
'        m_BinStr = m_BinStr + ECIDFuse.Category(idx2).Write.BitstrM
'        m_BinStr = m_BinStr + ECIDFuse.Category(idx3).Write.BitstrM
'
'        '< Convert the joint binary string to user Hexadecimal code >
'        ''''was m_hexStr = auto_Binary2Hex(StrReverse(m_BinStr))
'        m_hexStr = auto_BinStr2HexStr(StrReverse(m_BinStr), m_hexlen)
        
        m_HexStr = auto_WaferData_to_HexECID(ECIDFuse.Category(idx0).Write.BitStrM, _
                                             ECIDFuse.Category(idx1).Write.BitStrM, _
                                             ECIDFuse.Category(idx2).Write.BitStrM, _
                                             ECIDFuse.Category(idx3).Write.BitStrM)

        
    Next m_Site

    Set auto_WaferData_to_HexECID_SiteAware = m_HexStr
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


''''It checks only for the bits(DEID) and reserved bits.
Public Function auto_ECID_SyntaxCheck_DEID(Optional Ecid_ProgBit_Str As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ECID_SyntaxCheck"
    
    Dim FirstWord As String, Lot2_6 As String
    Dim Dec_val As Variant
    Dim Flag1 As Long, FlagX As Long, FlagY As Long
    Dim i As Long
    Dim Sum1 As Long
    Dim PrintStr As String
    Dim WaferVal As Long
    Dim n As Long
    Dim idx As Long

    Dim LotID_FirstChar As New SiteVariant
    Dim LotID_2_to_6Char As New SiteVariant
    Dim ProgBit_Sum As New SiteVariant
    Dim WaferID_Syntax As New SiteVariant
    Dim XY_Coor_Range As New SiteVariant
    
    Dim andFlag1 As Long
    Dim tmpName As String
    Dim m_result As Long
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant

    ''''20161018 for CRC
    Dim m_crchexStr As String
    Dim m_stage As String
    Dim m_valStr As String
    Dim m_bitStrM As String
    
    Dim m_bitstrL As String
    Dim m_bitwidth As Long
    Dim m_defval As String
    Dim m_defreal As String
    
    Dim ss As Variant
    
    ''''201811XX
    Dim m_customUnit As String
    Dim mSL_bitSum As New SiteLong
    Dim mSV_value As New SiteVariant
    Dim mSV_decimal As New SiteVariant
    Dim mSV_bitStrM As New SiteVariant
    Dim mSV_hexStr As New SiteVariant
    Dim mSL_valueSum As New SiteLong

    ''''-------------------------------------------------------------
    Dim m_siteVar As String
    Dim m_siteValue As Long
    m_siteVar = "ECIDChk_Var"
    m_siteValue = TheExec.sites(ss).SiteVariableValue(m_siteVar)
    ''''-------------------------------------------------------------

    'Initialization
    andFlag1 = 0
    Sum1 = 0
    WaferVal = 0
    Flag1 = 1
   
    Dim mStr_ECID_ValStr As String
    Dim mStr_ECID_ProgBit As String
    Dim mL_MSBbit As Long
    Dim mL_LSBbit As Long
    
    '' For LotID, waferID, XCoord, YCoord
    For Each site In TheExec.sites
        ss = TheExec.sites.SiteNumber
        With ECIDFuse.Category(ECIDIndex("Lot_ID"))
            mL_MSBbit = .MSBbit
            mL_LSBbit = .LSBbit
            mStr_ECID_ValStr = .Read.ValStr(site)
'            mStr_ECID_ProgBit = .Read.BitstrM
'            If (mL_MSBbit < mL_LSBbit) Then mStr_ECID_ProgBit = .Read.BitStrL
            
        End With
        
        'Step1. The first letter of LotID has to be numeric [A-Z].
        FirstWord = Mid(mStr_ECID_ValStr, 1, 1)
        If (FirstWord = "") Then
            Flag1 = 0   'Fail
        Else
            Dec_val = Asc(LCase(FirstWord))
            If Dec_val < 97 Or Dec_val > 122 Then   'a=97 and z=122 in ANSI character
                Flag1 = 0   'Fail
                LotID_FirstChar(ss) = "First Character of Lot ID is not [A-Z] (site :" + CStr(ss) + ")."
                TheExec.Datalog.WriteComment LotID_FirstChar(ss)
            Else
                Flag1 = 1 'Pass
            End If
        End If
        andFlag1 = Flag1
        m_catename = "LotID_1st_Char"
        TheExec.Flow.TestLimit Flag1, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename ''''tname:="First Lot Char"  'BurstResult=1:Pass
        
        'Step2. The 2nd~6th letters of LotID have to be either [0-9] or [A-Z].
        If (Len(mStr_ECID_ValStr) <> EcidCharPerLotId) Then
            Flag1 = 0   'Fail
        Else
            For i = 2 To EcidCharPerLotId  ''''EcidCharPerLotId=6
                ''Lot2_6 = Mid(HramLotId, i, 1)
                Lot2_6 = Mid(mStr_ECID_ValStr, i, 1)
                Dec_val = Asc(LCase(Lot2_6))
                If Dec_val < 97 Or Dec_val > 122 Then   'a=97 and z=122 in ANSI character
                    If Dec_val < 48 Or Dec_val > 57 Then   ''0'=48 and '9'=57 in ANSI character
                        Flag1 = 0   'Fail
                        LotID_2_to_6Char(ss) = "Second-to-Sixth Characters of Lot ID are not [A-Z] or [0-9] (site :" + CStr(ss) + ")."
                        TheExec.Datalog.WriteComment LotID_2_to_6Char(ss)
                        Exit For
                    Else
                        Flag1 = 1 'Pass
                    End If
                Else
                    Flag1 = 1 'Pass
                End If
            Next i
        End If
        andFlag1 = andFlag1 And Flag1
        m_catename = "LotID_2to6_Char"
        TheExec.Flow.TestLimit Flag1, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename ''''tname:="2nd to 6th Lot Char"  'BurstResult=1:Pass

        ''''20160118 New
        If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
            If (UCase(LotID) = UCase(mStr_ECID_ValStr)) Then
                ''''Pass
                Flag1 = 1
            Else
                ''''Fail
                Flag1 = 0
            End If
            andFlag1 = andFlag1 And Flag1
            TheExec.Flow.TestLimit Flag1, 1, 1, Tname:="Prober_" + UCase(LotID) + "_vs_DUT_" + UCase(mStr_ECID_ValStr)
        End If
        
        'Step3. Check the summation for bit 6 to bit52  > 7
        If (mStr_ECID_ProgBit <> "") Then
        Else
            Dim tmp As String:: tmp = ""
            For i = 0 To 3
                With ECIDFuse.Category(i)
                    If (mL_MSBbit < mL_LSBbit) Then
                         tmp = .Read.BitStrL
                    Else
                        tmp = .Read.BitStrM
                    End If
                    mStr_ECID_ProgBit = mStr_ECID_ProgBit + tmp
                End With
            Next i
        End If
        
        m_lolmt = 7
        m_hilmt = Len(mStr_ECID_ProgBit)
        m_catename = "ECID_" + CStr(m_hilmt) + "bits"
        Sum1 = 0
        For i = 0 To m_hilmt - 1
            Sum1 = Sum1 + CLng(Mid(mStr_ECID_ProgBit, m_hilmt - i, 1))
        Next i
        
        If Sum1 <= m_lolmt Then
            ProgBit_Sum(ss) = "Summation of bit0 to bit " & (m_hilmt - 1) & " data is wrong. Because it's <=" + CStr(m_lolmt) + " (site :" & CStr(ss) & ")."
            TheExec.Datalog.WriteComment ProgBit_Sum(ss)
            Flag1 = 0   'Fail
        Else
            Flag1 = 1 'Pass
        End If
        andFlag1 = andFlag1 And Flag1
        TheExec.Flow.TestLimit Sum1, m_lolmt, m_hilmt, tlSignGreaterEqual, tlSignLessEqual, Tname:="Sum_" + m_catename '''tname:="LotID summation"  'BurstResult=1:Pass
    
        'Step4. Check  Wafer Syntax
        idx = ECIDIndex("Wafer_ID")
        m_catename = ECIDFuse.Category(idx).Name
        m_lolmt = ECIDFuse.Category(idx).LoLMT
        m_hilmt = ECIDFuse.Category(idx).HiLMT
        WaferVal = ECIDFuse.Category(idx).Read.Decimal(ss)
        If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
            ''''20160118, New
            m_lolmt = WaferID
            m_hilmt = WaferID
        End If
        If WaferVal < m_lolmt Or WaferVal > m_hilmt Then
            Flag1 = 0   'Fail
            WaferID_Syntax(ss) = "WaferID is wrong. Because it's not in  [1-25] (site : " + CStr(ss) + ")."
            TheExec.Datalog.WriteComment WaferID_Syntax(ss)
        Else
            Flag1 = 1   'Pass
        End If
        andFlag1 = andFlag1 And Flag1
        TheExec.Flow.TestLimit WaferVal, m_lolmt, m_hilmt, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename
    
        'Step5. Check  X,Y Syntax
        Dim m_resultX As Long
        Dim m_resultY As Long
        FlagX = 0: FlagY = 0
        m_resultX = ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.Decimal(ss)
        m_resultY = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.Decimal(ss)
        If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
            ''''20160118 New
            XCOORD_LoLMT = XCoord(ss)
            XCOORD_HiLMT = XCoord(ss)
            ''''-----------------------
            YCOORD_LoLMT = YCoord(ss)
            YCOORD_HiLMT = YCoord(ss)
        End If
        If ((XCOORD_LoLMT <= m_resultX) And (m_resultX <= XCOORD_HiLMT)) Then FlagX = 1   'Pass
        If ((YCOORD_LoLMT <= m_resultY) And (m_resultY <= YCOORD_HiLMT)) Then FlagY = 1   'Pass
        If FlagX = 1 And FlagY = 1 Then
            Flag1 = 1
        Else
            Flag1 = 0  'Fail
            XY_Coor_Range(ss) = "X or  Y Coordinates (" & m_resultX & "," & m_resultY & ") are wrong (Site : " & CStr(ss) & ")."
            TheExec.Datalog.WriteComment XY_Coor_Range(ss)
        End If
        andFlag1 = andFlag1 And Flag1
        ''m_catename = ECIDFuse.Category(ECIDIndex("X_Coordinate")).Name
        m_catename = "X_Coordinate"
        TheExec.Flow.TestLimit m_resultX, XCOORD_LoLMT, XCOORD_HiLMT, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename  'BurstResult=1:Pass
    
        ''m_catename = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Name
        m_catename = "Y_Coordinate"
        TheExec.Flow.TestLimit m_resultY, YCOORD_LoLMT, YCOORD_HiLMT, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename  'BurstResult=1:Pass
    Next site
    
    Dim m_testFlag As Long
    Dim m_bitsum As Long
    'Step6, Check Reserved Bits
    For i = 0 To UBound(ECIDFuse.Category)
'        m_algorithm = LCase(ECIDFuse.Category(i).Algorithm)
'        m_catename = ECIDFuse.Category(i).Name
'        m_stage = LCase(ECIDFuse.Category(i).Stage)
'        m_result = ECIDFuse.Category(i).Read.Decimal(ss)
'        m_bitSum = ECIDFuse.Category(i).Read.BitSummation(ss)
'        m_hilmt = ECIDFuse.Category(i).HiLMT
'        m_lolmt = ECIDFuse.Category(i).LoLMT
'        m_valStr = ECIDFuse.Category(i).Read.ValStr(ss)
'        m_bitstrM = ECIDFuse.Category(i).Read.BitstrM(ss) ''''20161013 update
        
        With ECIDFuse.Category(i)
            m_stage = LCase(.Stage)
            m_catename = .Name
            m_algorithm = LCase(.algorithm)
            'm_value = .Read.Value(Site)
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
            m_bitwidth = .BitWidth
            m_defval = .DefaultValue
            'm_resolution = .Resoultion
            m_defreal = LCase(.Default_Real)
            ''''-----------------------------------
            mSL_bitSum = .Read.BitSummation
            mSV_decimal = .Read.Decimal
            mSV_bitStrM = .Read.BitStrM
            mSV_hexStr = .Read.HexStr
            mSV_value = .Read.Value
            ''''-----------------------------------
        End With
        
        tmpName = m_catename
        m_testFlag = 0
        Flag1 = 1
        tmpName = Replace(tmpName, " ", "_") ''''20151028, benefit for the script

        If ((m_algorithm = "lotid") Or (m_algorithm = "numeric")) Then
            ''''Has been checked on the above statement.
            m_testFlag = 0

        ''''' 20161003 ADDC RC
        ElseIf (m_algorithm = "crc") Then ''''20160924 add
            m_testFlag = 0
            ''''20161014 ECID CRC always fuse at CP1 stage !
            ''''So it will have no this condition "gB_eFuse_Disable_ChkLMT_Flag=True"
            For Each site In TheExec.sites
                If (gS_ECID_CRC_PgmFlow = UCase("DEID")) Then ''''<NOTICE> 20170815 update
                    m_crchexStr = UCase(gS_ECID_CRC_HexStr(ss)) ''''it comes from the calculation from CRC calcBits
                    m_valStr = auto_BinStr2HexStr(m_bitStrM, 4)
    
                    tmpName = tmpName + "_" + m_crchexStr
                    If (UCase(m_valStr) = m_crchexStr) Then
                        ''''Pass for CRC correct
                        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=tmpName
                    Else
                        '''Fail
                        TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=tmpName
                    End If
                End If
            Next site

        ElseIf (m_algorithm = LCase("DEID")) Then ''''skip the last one "ECID_DEID"
            m_testFlag = 0

        ElseIf (m_algorithm Like "*reserve*") Then
            m_result = m_bitsum
            If (m_result < m_lolmt Or m_result > m_hilmt) Then
                TheExec.Datalog.WriteComment m_catename + " failed to test limits."
                'TheExec.Datalog.WriteComment "Site(" + CStr(ss) + ")::" + m_catename + " failed to test limits."
                Flag1 = 0
            End If
            andFlag1 = andFlag1 And Flag1
            m_testFlag = 1
        Else
''''            ''''<NOTICE> 20151221 Add
''''            ''''for others, it should be NOT blown so the limit is zero in the very 1st time
''''            If (m_siteValue = 1) Then
''''                m_lolmt = 0
''''                m_hilmt = 0
''''            End If
''''            If (m_result < m_lolmt Or m_result > m_hilmt) Then
''''                TheExec.Datalog.WriteComment "Site(" + CStr(ss) + ")::" + m_catename + " failed to test limits."
''''                Flag1 = 0
''''            End If
''''            andFlag1 = andFlag1 And Flag1
''''            m_testflag = 1

            ''''20160907 update, it could cause the failure on retest run if the nonDEID does NOT be blown in the 1st run flow.
            m_testFlag = 0 ''''bypass the non-DEID category syntax check here
        End If

        If (m_testFlag = 1) Then
            TheExec.Flow.TestLimit m_result, m_lolmt, m_hilmt, tlSignGreaterEqual, tlSignLessEqual, Tname:=tmpName
        End If
    Next i
    

'
'    Dim myStr As String
'    myStr = ECIDFuse.Category(ECIDIndex("Lot_ID")).Read.ValStr(ss)
'
'    'Step1. The first letter of LotID has to be numeric [A-Z].
'    FirstWord = Mid(myStr, 1, 1)
'    If (FirstWord = "") Then
'        Flag1 = 0   'Fail
'    Else
'        Dec_Val = Asc(LCase(FirstWord))
'        If Dec_Val < 97 Or Dec_Val > 122 Then   'a=97 and z=122 in ANSI character
'            Flag1 = 0   'Fail
'            LotID_FirstChar(ss) = "First Character of Lot ID is not [A-Z] (site :" + CStr(ss) + ")."
'            TheExec.Datalog.WriteComment LotID_FirstChar(ss)
'        Else
'            Flag1 = 1 'Pass
'        End If
'    End If
'    andFlag1 = Flag1
'    m_catename = "LotID_1st_Char"
'    TheExec.Flow.TestLimit Flag1, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename ''''tname:="First Lot Char"  'BurstResult=1:Pass
'
'    'Step2. The 2nd~6th letters of LotID have to be either [0-9] or [A-Z].
'    If (Len(myStr) <> EcidCharPerLotId) Then
'        Flag1 = 0   'Fail
'    Else
'        For i = 2 To EcidCharPerLotId  ''''EcidCharPerLotId=6
'            ''Lot2_6 = Mid(HramLotId, i, 1)
'            Lot2_6 = Mid(myStr, i, 1)
'            Dec_Val = Asc(LCase(Lot2_6))
'            If Dec_Val < 97 Or Dec_Val > 122 Then   'a=97 and z=122 in ANSI character
'                If Dec_Val < 48 Or Dec_Val > 57 Then   ''0'=48 and '9'=57 in ANSI character
'                    Flag1 = 0   'Fail
'                    LotID_2_to_6Char(ss) = "Second-to-Sixth Characters of Lot ID are not [A-Z] or [0-9] (site :" + CStr(ss) + ")."
'                    TheExec.Datalog.WriteComment LotID_2_to_6Char(ss)
'                    Exit For
'                Else
'                    Flag1 = 1 'Pass
'                End If
'            Else
'                Flag1 = 1 'Pass
'            End If
'        Next i
'    End If
'    andFlag1 = andFlag1 And Flag1
'    m_catename = "LotID_2to6_Char"
'    TheExec.Flow.TestLimit Flag1, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename ''''tname:="2nd to 6th Lot Char"  'BurstResult=1:Pass
'
'    ''''20160118 New
'    If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
'        If (UCase(LotID) = UCase(myStr)) Then
'            ''''Pass
'            Flag1 = 1
'        Else
'            ''''Fail
'            Flag1 = 0
'        End If
'        andFlag1 = andFlag1 And Flag1
'        TheExec.Flow.TestLimit Flag1, 1, 1, Tname:="Prober_" + UCase(LotID) + "_vs_DUT_" + UCase(myStr)
'    End If
'
'    'Step3. Check the summation for bit 6 to bit52  > 7
'    m_lolmt = 7
'    m_hilmt = Len(Ecid_ProgBit_Str)
'    m_catename = "ECID_" + CStr(m_hilmt) + "bits"
'    Sum1 = 0
'    For i = 0 To m_hilmt - 1
'        Sum1 = Sum1 + CLng(Mid(Ecid_ProgBit_Str, m_hilmt - i, 1))
'    Next i
'
'    If Sum1 <= m_lolmt Then
'        ProgBit_Sum(ss) = "Summation of bit0 to bit " & (m_hilmt - 1) & " data is wrong. Because it's <=" + CStr(m_lolmt) + " (site :" & CStr(ss) & ")."
'        TheExec.Datalog.WriteComment ProgBit_Sum(ss)
'        Flag1 = 0   'Fail
'    Else
'        Flag1 = 1 'Pass
'    End If
'    andFlag1 = andFlag1 And Flag1
'    TheExec.Flow.TestLimit Sum1, m_lolmt, m_hilmt, tlSignGreaterEqual, tlSignLessEqual, Tname:="Sum_" + m_catename '''tname:="LotID summation"  'BurstResult=1:Pass
'
'
'    'Step4. Check  Wafer Syntax
'    idx = ECIDIndex("Wafer_ID")
'    m_catename = ECIDFuse.Category(idx).Name
'    m_lolmt = ECIDFuse.Category(idx).LoLMT
'    m_hilmt = ECIDFuse.Category(idx).HiLMT
'    WaferVal = ECIDFuse.Category(idx).Read.Decimal(ss)
'    If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
'        ''''20160118, New
'        m_lolmt = WaferID
'        m_hilmt = WaferID
'    End If
'    If WaferVal < m_lolmt Or WaferVal > m_hilmt Then
'        Flag1 = 0   'Fail
'        WaferID_Syntax(ss) = "WaferID is wrong. Because it's not in  [1-25] (site : " + CStr(ss) + ")."
'        TheExec.Datalog.WriteComment WaferID_Syntax(ss)
'    Else
'        Flag1 = 1   'Pass
'    End If
'    andFlag1 = andFlag1 And Flag1
'    TheExec.Flow.TestLimit WaferVal, m_lolmt, m_hilmt, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename
'
'    'Step5. Check  X,Y Syntax
'    Dim m_resultX As Long
'    Dim m_resultY As Long
'    FlagX = 0: FlagY = 0
'    m_resultX = ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.Decimal(ss)
'    m_resultY = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.Decimal(ss)
'    If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
'        ''''20160118 New
'        XCOORD_LoLMT = XCoord(ss)
'        XCOORD_HiLMT = XCoord(ss)
'        ''''-----------------------
'        YCOORD_LoLMT = YCoord(ss)
'        YCOORD_HiLMT = YCoord(ss)
'    End If
'    If ((XCOORD_LoLMT <= m_resultX) And (m_resultX <= XCOORD_HiLMT)) Then FlagX = 1   'Pass
'    If ((YCOORD_LoLMT <= m_resultY) And (m_resultY <= YCOORD_HiLMT)) Then FlagY = 1   'Pass
'    If FlagX = 1 And FlagY = 1 Then
'        Flag1 = 1
'    Else
'        Flag1 = 0  'Fail
'        XY_Coor_Range(ss) = "X or  Y Coordinates (" & m_resultX & "," & m_resultY & ") are wrong (Site : " & CStr(ss) & ")."
'        TheExec.Datalog.WriteComment XY_Coor_Range(ss)
'    End If
'    andFlag1 = andFlag1 And Flag1
'    ''m_catename = ECIDFuse.Category(ECIDIndex("X_Coordinate")).Name
'    m_catename = "X_Coordinate"
'    TheExec.Flow.TestLimit m_resultX, XCOORD_LoLMT, XCOORD_HiLMT, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename  'BurstResult=1:Pass
'
'    ''m_catename = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Name
'    m_catename = "Y_Coordinate"
'    TheExec.Flow.TestLimit m_resultY, YCOORD_LoLMT, YCOORD_HiLMT, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename  'BurstResult=1:Pass
'
'
'    Dim m_testFlag As Long
'    Dim m_bitSum As Long
'    'Step6, Check Reserved Bits
'    For i = 0 To UBound(ECIDFuse.Category)
'        m_algorithm = LCase(ECIDFuse.Category(i).Algorithm)
'        m_catename = ECIDFuse.Category(i).Name
'        m_stage = LCase(ECIDFuse.Category(i).Stage)
'        m_result = ECIDFuse.Category(i).Read.Decimal(ss)
'        m_bitSum = ECIDFuse.Category(i).Read.BitSummation(ss)
'        m_hilmt = ECIDFuse.Category(i).HiLMT
'        m_lolmt = ECIDFuse.Category(i).LoLMT
'        m_valStr = ECIDFuse.Category(i).Read.ValStr(ss)
'        m_bitstrM = ECIDFuse.Category(i).Read.BitstrM(ss) ''''20161013 update
'        tmpName = m_catename
'        m_testFlag = 0
'        Flag1 = 1
'        tmpName = Replace(tmpName, " ", "_") ''''20151028, benefit for the script
'
'        If ((m_algorithm = "lotid") Or (m_algorithm = "numeric")) Then
'            ''''Has been checked on the above statement.
'            m_testFlag = 0
'
'        ''''' 20161003 ADDC RC
'        ElseIf (m_algorithm = "crc") Then ''''20160924 add
'            m_testFlag = 0
'            ''''20161014 ECID CRC always fuse at CP1 stage !
'            ''''So it will have no this condition "gB_eFuse_Disable_ChkLMT_Flag=True"
'
'            If (gS_ECID_CRC_PgmFlow = UCase("DEID")) Then ''''<NOTICE> 20170815 update
'                m_crchexStr = UCase(gS_ECID_CRC_HexStr(ss)) ''''it comes from the calculation from CRC calcBits
'                m_valStr = auto_BinStr2HexStr(m_bitstrM, 4)
'
'                tmpName = tmpName + "_" + m_crchexStr
'                If (UCase(m_valStr) = m_crchexStr) Then
'                    ''''Pass for CRC correct
'                    TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=tmpName
'                Else
'                    '''Fail
'                    TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=tmpName
'                End If
'            End If
'
'        ElseIf (m_algorithm = LCase("DEID")) Then ''''skip the last one "ECID_DEID"
'            m_testFlag = 0
'
'        ElseIf (m_algorithm Like "*reserve*") Then
'            m_result = m_bitSum
'            If (m_result < m_lolmt Or m_result > m_hilmt) Then
'                TheExec.Datalog.WriteComment "Site(" + CStr(ss) + ")::" + m_catename + " failed to test limits."
'                Flag1 = 0
'            End If
'            andFlag1 = andFlag1 And Flag1
'            m_testFlag = 1
'        Else
'''''            ''''<NOTICE> 20151221 Add
'''''            ''''for others, it should be NOT blown so the limit is zero in the very 1st time
'''''            If (m_siteValue = 1) Then
'''''                m_lolmt = 0
'''''                m_hilmt = 0
'''''            End If
'''''            If (m_result < m_lolmt Or m_result > m_hilmt) Then
'''''                TheExec.Datalog.WriteComment "Site(" + CStr(ss) + ")::" + m_catename + " failed to test limits."
'''''                Flag1 = 0
'''''            End If
'''''            andFlag1 = andFlag1 And Flag1
'''''            m_testflag = 1
'
'            ''''20160907 update, it could cause the failure on retest run if the nonDEID does NOT be blown in the 1st run flow.
'            m_testFlag = 0 ''''bypass the non-DEID category syntax check here
'        End If
'
'        If (m_testFlag = 1) Then
'            TheExec.Flow.TestLimit m_result, m_lolmt, m_hilmt, tlSignGreaterEqual, tlSignLessEqual, Tname:=tmpName
'        End If
'    Next i
'
'    ''''using in outside to check if all (Flag1)s' and-result is '1'(pass) or '0'(fail)
'    '**'auto_Chk_ECID_Content_DEID andFlag1

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function


''''It checks only for the bits(DEID) and reserved bits.
Public Function auto_ECID_SyntaxCheck_All(Optional Ecid_ProgBit_Str As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ECID_SyntaxCheck"
    
    Dim FirstWord As String, Lot2_6 As String
    Dim Dec_val As Variant
    Dim Flag1 As Long, FlagX As Long, FlagY As Long
    Dim i As Long
    Dim Sum1 As Long
    Dim PrintStr As String
    Dim WaferVal As Long
    Dim n As Long
    Dim idx As Long

    Dim LotID_FirstChar As New SiteVariant
    Dim LotID_2_to_6Char As New SiteVariant
    Dim ProgBit_Sum As New SiteVariant
    Dim WaferID_Syntax As New SiteVariant
    Dim XY_Coor_Range As New SiteVariant
    
    Dim andFlag1 As Long
    Dim tmpName As String
    Dim m_result As Long
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant

    ''''20161018 for CRC
    Dim m_crchexStr As String
    Dim m_stage As String
    Dim m_valStr As String
    Dim m_bitStrM As String
    Dim m_testValue As Variant
    
    Dim m_bitstrL As String
    'Dim m_BitWidth As String
    Dim m_defval As String
    Dim m_defreal As String
    Dim m_testFlag As Integer
    Dim m_bitwidth As Long
    Dim m_HexStr As String
    
    Dim ss As Variant
    
    ''''201811XX
    Dim m_customUnit As String
    Dim mSL_bitSum As New SiteLong
    Dim mSV_value As New SiteVariant
    Dim mSV_decimal As New SiteVariant
    Dim mSV_bitStrM As New SiteVariant
    Dim mSV_hexStr As New SiteVariant
    Dim mSL_valueSum As New SiteLong

    ''''-------------------------------------------------------------
    Dim m_siteVar As String
    Dim m_siteValue As Long
    m_siteVar = "ECIDChk_Var"
    m_siteValue = TheExec.sites(ss).SiteVariableValue(m_siteVar)
    ''''-------------------------------------------------------------

    'Initialization
    andFlag1 = 0
    Sum1 = 0
    WaferVal = 0
    Flag1 = 1
   
    Dim mStr_ECID_ValStr As String
    Dim mStr_ECID_ProgBit As String
    Dim mL_MSBbit As Long
    Dim mL_LSBbit As Long
    Dim tmpName0 As String
    
    '' For LotID, waferID, XCoord, YCoord
    For Each site In TheExec.sites
        ss = TheExec.sites.SiteNumber
        With ECIDFuse.Category(ECIDIndex("Lot_ID"))
            mL_MSBbit = .MSBbit
            mL_LSBbit = .LSBbit
            mStr_ECID_ValStr = .Read.ValStr(site)
'            mStr_ECID_ProgBit = .Read.BitstrM
'            If (mL_MSBbit < mL_LSBbit) Then mStr_ECID_ProgBit = .Read.BitStrL
            
        End With
        
        'Step1. The first letter of LotID has to be numeric [A-Z].
        FirstWord = Mid(mStr_ECID_ValStr, 1, 1)
        If (FirstWord = "") Then
            Flag1 = 0   'Fail
        Else
            Dec_val = Asc(LCase(FirstWord))
            If Dec_val < 97 Or Dec_val > 122 Then   'a=97 and z=122 in ANSI character
                Flag1 = 0   'Fail
                LotID_FirstChar(ss) = "First Character of Lot ID is not [A-Z] (site :" + CStr(ss) + ")."
                TheExec.Datalog.WriteComment LotID_FirstChar(ss)
            Else
                Flag1 = 1 'Pass
            End If
        End If
        andFlag1 = Flag1
        m_catename = "LotID_1st_Char"
        TheExec.Flow.TestLimit Flag1, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename ''''tname:="First Lot Char"  'BurstResult=1:Pass
        
        'Step2. The 2nd~6th letters of LotID have to be either [0-9] or [A-Z].
        If (Len(mStr_ECID_ValStr) <> EcidCharPerLotId) Then
            Flag1 = 0   'Fail
        Else
            For i = 2 To EcidCharPerLotId  ''''EcidCharPerLotId=6
                ''Lot2_6 = Mid(HramLotId, i, 1)
                Lot2_6 = Mid(mStr_ECID_ValStr, i, 1)
                Dec_val = Asc(LCase(Lot2_6))
                If Dec_val < 97 Or Dec_val > 122 Then   'a=97 and z=122 in ANSI character
                    If Dec_val < 48 Or Dec_val > 57 Then   ''0'=48 and '9'=57 in ANSI character
                        Flag1 = 0   'Fail
                        LotID_2_to_6Char(ss) = "Second-to-Sixth Characters of Lot ID are not [A-Z] or [0-9] (site :" + CStr(ss) + ")."
                        TheExec.Datalog.WriteComment LotID_2_to_6Char(ss)
                        Exit For
                    Else
                        Flag1 = 1 'Pass
                    End If
                Else
                    Flag1 = 1 'Pass
                End If
            Next i
        End If
        andFlag1 = andFlag1 And Flag1
        m_catename = "LotID_2to6_Char"
        TheExec.Flow.TestLimit Flag1, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename ''''tname:="2nd to 6th Lot Char"  'BurstResult=1:Pass

        ''''20160118 New
        If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
            If (UCase(LotID) = UCase(mStr_ECID_ValStr)) Then
                ''''Pass
                Flag1 = 1
            Else
                ''''Fail
                Flag1 = 0
            End If
            andFlag1 = andFlag1 And Flag1
            TheExec.Flow.TestLimit Flag1, 1, 1, Tname:="Prober_" + UCase(LotID) + "_vs_DUT_" + UCase(mStr_ECID_ValStr)
        End If
        
        'Step3. Check the summation for bit 6 to bit52  > 7
        If (mStr_ECID_ProgBit <> "") Then
        Else
            Dim tmp As String:: tmp = ""
            For i = 0 To UBound(ECIDFuse.Category) - 1
                With ECIDFuse.Category(i)
                    If (mL_MSBbit < mL_LSBbit) Then
                         tmp = .Read.BitStrL
                    Else
                        tmp = .Read.BitStrM
                    End If
                    mStr_ECID_ProgBit = mStr_ECID_ProgBit + tmp
                End With
            Next i
        End If
        m_lolmt = 7
        m_hilmt = Len(mStr_ECID_ProgBit)
        m_catename = "ECID_" + CStr(m_hilmt) + "bits"
        Sum1 = 0
        For i = 0 To m_hilmt - 1
            Sum1 = Sum1 + CLng(Mid(mStr_ECID_ProgBit, m_hilmt - i, 1))
        Next i
        
        If Sum1 <= m_lolmt Then
            ProgBit_Sum(ss) = "Summation of bit0 to bit " & (m_hilmt - 1) & " data is wrong. Because it's <=" + CStr(m_lolmt) + " (site :" & CStr(ss) & ")."
            TheExec.Datalog.WriteComment ProgBit_Sum(ss)
            Flag1 = 0   'Fail
        Else
            Flag1 = 1 'Pass
        End If
        andFlag1 = andFlag1 And Flag1
        TheExec.Flow.TestLimit Sum1, m_lolmt, m_hilmt, tlSignGreaterEqual, tlSignLessEqual, Tname:="Sum_" + m_catename '''tname:="LotID summation"  'BurstResult=1:Pass
    
        'Step4. Check  Wafer Syntax
        idx = ECIDIndex("Wafer_ID")
        m_catename = ECIDFuse.Category(idx).Name
        m_lolmt = ECIDFuse.Category(idx).LoLMT
        m_hilmt = ECIDFuse.Category(idx).HiLMT
        WaferVal = ECIDFuse.Category(idx).Read.Decimal(ss)
        If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
            ''''20160118, New
            m_lolmt = WaferID
            m_hilmt = WaferID
        End If
        If WaferVal < m_lolmt Or WaferVal > m_hilmt Then
            Flag1 = 0   'Fail
            WaferID_Syntax(ss) = "WaferID is wrong. Because it's not in  [1-25] (site : " + CStr(ss) + ")."
            TheExec.Datalog.WriteComment WaferID_Syntax(ss)
        Else
            Flag1 = 1   'Pass
        End If
        andFlag1 = andFlag1 And Flag1
        TheExec.Flow.TestLimit WaferVal, m_lolmt, m_hilmt, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename
    
        'Step5. Check  X,Y Syntax
        Dim m_resultX As Long
        Dim m_resultY As Long
        FlagX = 0: FlagY = 0
        m_resultX = ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.Decimal(ss)
        m_resultY = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.Decimal(ss)
        If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
            ''''20160118 New
            XCOORD_LoLMT = XCoord(ss)
            XCOORD_HiLMT = XCoord(ss)
            ''''-----------------------
            YCOORD_LoLMT = YCoord(ss)
            YCOORD_HiLMT = YCoord(ss)
        End If
        If ((XCOORD_LoLMT <= m_resultX) And (m_resultX <= XCOORD_HiLMT)) Then FlagX = 1   'Pass
        If ((YCOORD_LoLMT <= m_resultY) And (m_resultY <= YCOORD_HiLMT)) Then FlagY = 1   'Pass
        If FlagX = 1 And FlagY = 1 Then
            Flag1 = 1
        Else
            Flag1 = 0  'Fail
            XY_Coor_Range(ss) = "X or  Y Coordinates (" & m_resultX & "," & m_resultY & ") are wrong (Site : " & CStr(ss) & ")."
            TheExec.Datalog.WriteComment XY_Coor_Range(ss)
        End If
        andFlag1 = andFlag1 And Flag1
        ''m_catename = ECIDFuse.Category(ECIDIndex("X_Coordinate")).Name
        m_catename = "X_Coordinate"
        TheExec.Flow.TestLimit m_resultX, XCOORD_LoLMT, XCOORD_HiLMT, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename  'BurstResult=1:Pass
    
        ''m_catename = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Name
        m_catename = "Y_Coordinate"
        TheExec.Flow.TestLimit m_resultY, YCOORD_LoLMT, YCOORD_HiLMT, tlSignGreaterEqual, tlSignLessEqual, Tname:=m_catename  'BurstResult=1:Pass
    Next site
    
    'Step6, Check others
    For i = 0 To UBound(ECIDFuse.Category) - 1 ''''skip "DEID"
        With ECIDFuse.Category(i)
            m_stage = LCase(.Stage)
            m_catename = .Name
            m_algorithm = LCase(.algorithm)
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
            m_bitwidth = .BitWidth
            m_defval = .DefaultValue
            m_defreal = LCase(.Default_Real)
            ''''-----------------------------------
            mSL_bitSum = .Read.BitSummation
            mSV_decimal = .Read.Decimal
            mSV_bitStrM = .Read.BitStrM
            mSV_hexStr = .Read.HexStr
            mSV_value = .Read.Value
            ''''-----------------------------------
        End With

        tmpName = m_catename
        m_testFlag = 0
        Flag1 = 1

        tmpName = Replace(tmpName, " ", "_") ''''20151028, benefit for the script

        If ((m_algorithm = "lotid") Or (m_algorithm = "numeric")) Then
            ''''Has been checked on the above statement.
            m_testFlag = 0

        '''''20161018 for CRC update, 20161013 update
        ElseIf (m_algorithm = "crc") Then
            m_testFlag = 0
            ''''20161014 ECID CRC always fuse at CP1 stage !
            ''''So it will have no this condition "gB_eFuse_Disable_ChkLMT_Flag=True"
            ''''<NOTICE> 20161018 update
            ''''Because nonDEID will check CRC again, but there is no Programming CRC,
            ''''So we just use gS_ECID_CRC_HexStr as the reference, not using Write.ValStr
'            For Each Site In TheExec.Sites
'                m_crchexStr = UCase(gS_ECID_CRC_HexStr(Site)) ''''it comes from the calculation from bit63 to bit0
'
'                m_valStr = auto_BinStr2HexStr(m_bitstrM, 4)
'
'                tmpName = tmpName + "_" + m_crchexStr
'                If (UCase(CStr(m_valStr)) = m_crchexStr) Then
'                    ''''Pass for CRC correct
'                    TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=tmpName
'                Else
'                    '''Fail
'                    Flag1 = 0
'                    TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=tmpName
'                End If
'            Next Site
'                andFlag1 = andFlag1 And Flag1

                    m_lolmt = 0
                    m_hilmt = 0
                    tmpName0 = tmpName
                    For Each site In TheExec.sites
                            'm_crchexStr = "0x" + auto_BinStr2HexStr(StrReverse(gS_ECID_Read_calcCRC_bitStrM(Site)), m_bitwidth / 4)
                            m_crchexStr = "0x" + auto_BinStr2HexStr((mSV_bitStrM(site)), m_bitwidth / 4)
                            tmpName = tmpName0 + "_" + m_crchexStr
                            ''''<NOTICE>
                            ''''mSV_hexStr  is the CRC HexStr of Read eFuse Category
                            ''''m_crchexStr is the CRC HexStr by the calculation of read bits.
                            If (UCase(mSV_hexStr) = UCase(m_crchexStr)) Then
                                ''''Pass
                                TheExec.Flow.TestLimit resultVal:=0, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=tmpName
                            Else
                                ''''Fail
                                TheExec.Flow.TestLimit resultVal:=1, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=tmpName
                            End If
                        Next site



        ElseIf (m_algorithm = LCase("DEID")) Then ''''skip the last one "ECID_DEID"
            m_testFlag = 0
        ElseIf (m_algorithm = "wat") Then
            ''''20151014 update
            m_lolmt = 0
            m_hilmt = CDbl(2 ^ m_bitwidth) - 1
            m_testFlag = 1

        Else
            m_testFlag = 1
        End If

        If (m_testFlag = 1) Then
            ''''----------------------------------------------------
            ''''20170911 New
            ''''----------------------------------------------------
            Call auto_eFuse_chkLoLimit("ECID", i, m_stage, m_lolmt)
            Call auto_eFuse_chkHiLimit("ECID", i, m_stage, m_hilmt)
            
             ''''20170811 update
            If (m_bitwidth >= 32) Then
                ''m_tsName = m_tsName + "_" + m_hexStr
                ''''process m_lolmt / m_hilmt to hex string with prefix '0x'
                m_lolmt = auto_Value2HexStr(m_lolmt, m_bitwidth)
                m_hilmt = auto_Value2HexStr(m_hilmt, m_bitwidth)
                
                ''''------------------------------------------
                ''''compare with lolmt, hilmt
                ''''m_testValue 0 means fail
                ''''m_testValue 1 means pass
                ''''------------------------------------------
                For Each site In TheExec.sites
                    m_HexStr = mSV_hexStr(site)
                    mSV_value(site) = auto_TestStringLimit(m_HexStr, CStr(m_lolmt), CStr(m_hilmt)) - 1
                Next site
                    ''''mSV_value=0: Pass, = -1 Fail
                m_lolmt = 0
                m_hilmt = 0
            Else
                ''''20160927 update the new logical methodology for the unexpected binary decode.
                If (auto_isHexString(CStr(m_lolmt)) = True) Then
                    ''''translate to double value
                    m_lolmt = auto_HexStr2Value(m_lolmt)
                Else
                    ''''doNothing, m_lolmt = m_lolmt
                End If

                If (auto_isHexString(CStr(m_hilmt)) = True) Then
                    ''''translate to double value
                    m_hilmt = auto_HexStr2Value(m_hilmt)
                Else
                    ''''doNothing, m_hilmt = m_hilmt
                End If
                'mSV_value = m_result
            End If
            
'            If (CDbl(m_testValue) < m_lolmt Or CDbl(m_testValue) > m_hilmt) Then
'                TheExec.Datalog.WriteComment m_catename + " failed to test limits."
'                'TheExec.Datalog.WriteComment "Site(" + CStr(ss) + ")::" + m_catename + " failed to test limits."
'                Flag1 = 0
'            End If
            andFlag1 = andFlag1 And Flag1
            'TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=tmpName
            'TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=0, hiVal:=0, Tname:=m_tsName
            'TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=tmpName, ScaleType:=m_scale, unit:=m_unitType, customUnit:=m_customUnit
            If (andFlag1) Then
                TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=tmpName
            End If
            'TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=tmpName

            
            'TheExec.Flow.TestLimit m_testValue, m_lolmt, m_hilmt, tlSignGreaterEqual, tlSignLessEqual, Tname:=tmpName
            ''''----------------------------------------------------
        End If
    Next i
    


Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Public Sub auto_eFuse_Print_Device_code()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Print_Device_code"

    Dim m_tmp_Binary As String
    Dim m_user_ECID As String
    Dim m_user_ECID_Len As Long
    
    For Each site In TheExec.sites
    
        With ECIDFuse
        m_tmp_Binary = .Category(ECIDIndex("Y_Coordinate")).Read.BitStrM(site) & _
                       .Category(ECIDIndex("X_Coordinate")).Read.BitStrM(site) & _
                       .Category(ECIDIndex("Wafer_ID")).Read.BitStrM(site) & _
                       .Category(ECIDIndex("Lot_ID")).Read.BitStrM(site)
        End With
    
        m_user_ECID_Len = Len(m_tmp_Binary)
        Do Until (Len(m_tmp_Binary) Mod 64 = 0)
            m_tmp_Binary = "0" + m_tmp_Binary
        Loop
        m_user_ECID = auto_Binary2Hex(m_tmp_Binary)
            
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "Block_1" + " : "
        TheExec.Datalog.WriteComment "ECID Hexadecimal code = " + m_user_ECID
    
            
    
        ''''20160926 set it as the Default per Jack's comment before.
        ''''20161021 update per Laba and C651 PE request
        ''''It has been done in auto_ReadWaferData() already
        ''''20161118 Per Jack's comment, for WLFT case, set PRR from ECID of DUT.
        If (gB_ReadWaferData_flag = False Or gS_JobName Like "*wlft*") Then
            ''''Write to PRR-Part_TEXT in STDF
            Call TheExec.Datalog.Setup.SetPRRPartInfo(tl_SelectSite, site, , m_user_ECID)
        End If
    
        Dim m_device_code As String
        m_device_code = ECIDFuse.Category(ECIDIndex("Lot_ID")).Read.ValStr(site) & "_W" _
                        & ECIDFuse.Category(ECIDIndex("Wafer_ID")).Read.Value(site) & "_X" _
                        & ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.Value(site) & "_Y" _
                        & ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.Value(site) & "_S" & site
                                
        TheExec.Datalog.WriteComment "DEVICE_CODE: " + m_device_code
        TheExec.Datalog.WriteComment ""
    Next

Exit Sub

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Sub Else Resume Next
End Sub

