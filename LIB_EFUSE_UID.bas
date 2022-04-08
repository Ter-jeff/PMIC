Attribute VB_Name = "LIB_EFUSE_UID"
Option Explicit

'''''---------------------------------------------------------------------------------------------------
''''' UID (AES)  Fuse
'''''---------------------------------------------------------------------------------------------------

Public Function Cal_UIDChkSum(StrArry() As String, Block1Sum As Double, Block2Sum As Double) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "Cal_UIDChkSum"
    
    Dim i As Long, j As Long
    Dim ss As Variant
    ''Dim BlockRightSum As Double, BlockLeftSum As Double
    Dim k As Long
    Dim Count As Long
    ss = TheExec.sites.SiteNumber
    
    'Initialization
    Block1Sum = 0: Block2Sum = 0
    ''BlockRightSum = Block1Sum: BlockLeftSum = Block2Sum
    Count = 0
    
    If (gS_EFuse_Orientation = "UP2DOWN") Then
        Count = 0
        For i = 0 To UIDRowPerBlock - 1 ''0...31
            For j = 1 To UIDBitsPerRow  ''1...32
                ''''only sum to gL_UIDCodeBitWidth
                Count = Count + 1
                If (Count > gL_UIDCodeBitWidth) Then Exit For
                
                Block1Sum = Block1Sum + CDbl(Mid(StrArry(i, ss), j, 1))
            Next j
        Next i
        
        Count = 0
        For i = UIDRowPerBlock To (UIDBlock * UIDRowPerBlock - 1) ''32...63
            For j = 1 To UIDBitsPerRow  ''1...32
                ''''only sum to gL_UIDCodeBitWidth
                Count = Count + 1
                If (Count > gL_UIDCodeBitWidth) Then Exit For
                
                Block2Sum = Block2Sum + CDbl(Mid(StrArry(i, ss), j, 1))
            Next j
        Next i
        
        TheExec.Datalog.WriteComment "The CheckSum in Block1 Up  (site " & ss & ") = " & Block1Sum & " (" & Block1Sum / gL_UIDCodeBitWidth & ")"
        TheExec.Datalog.WriteComment "The CheckSum in Block2 Down(site " & ss & ") = " & Block2Sum & " (" & Block2Sum / gL_UIDCodeBitWidth & ")"

    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        Count = 0
        For i = 0 To UIDRowPerBlock - 1 ''0...63
            For j = 1 To UIDBitsPerRow  ''1...16
                ''''only sum to gL_UIDCodeBitWidth
                Count = Count + 1
                If (Count > gL_UIDCodeBitWidth) Then Exit For
                
                Block1Sum = Block1Sum + Mid(StrArry(i, ss), j, 1)
                Block2Sum = Block2Sum + Mid(StrArry(i, ss), j + UIDBitsPerRow, 1)
            Next j
        Next i

        TheExec.Datalog.WriteComment "The CheckSum in Block1 Right(site " & ss & ") = " & Block1Sum & " (" & Block1Sum / gL_UIDCodeBitWidth & ")"
        TheExec.Datalog.WriteComment "The CheckSum in Block2 Left (site " & ss & ") = " & Block2Sum & " (" & Block2Sum / gL_UIDCodeBitWidth & ")"

    ElseIf (gS_EFuse_Orientation = "SingleUp") Then
        Count = 0
        For i = 0 To UIDRowPerBlock - 1 ''0...31
            For j = 1 To UIDBitsPerRow  ''1...32
                ''''only sum to gL_UIDCodeBitWidth
                Count = Count + 1
                If (Count > gL_UIDCodeBitWidth) Then Exit For
                
                Block1Sum = Block1Sum + CDbl(Mid(StrArry(i, ss), j, 1))
            Next j
        Next i

        TheExec.Datalog.WriteComment "The CheckSum in Block (site " & ss & ") = " & Block1Sum & " (" & Block1Sum / gL_UIDCodeBitWidth & ")"

    ''''the below is reserved
    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UIDCompare_DoubleBit_PgmBit(DoubleBitArray() As Long, eFuse_Pgm_Bit() As Long, FailCnt As SiteLong)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UIDCompare_DoubleBit_PgmBit"
    
    Dim ss As Variant
    Dim k As Long, j As Long, i As Long
    
    ss = TheExec.sites.SiteNumber
    FailCnt(ss) = 0
    
    If (gS_EFuse_Orientation = "UP2DOWN") Then
 
        For i = 0 To (UIDBitsPerBlockUsed - 1) ''''UIDBitsPerBlockUsed=(UIDTotalBits / 2) in this case
            ''''Up-Side
            If (DoubleBitArray(i) <> eFuse_Pgm_Bit(i)) Then
                FailCnt(ss) = FailCnt(ss) + 1
            End If
            ''''Down-Side
            If (DoubleBitArray(i) <> eFuse_Pgm_Bit(i + UIDBitsPerBlockUsed)) Then
                FailCnt(ss) = FailCnt(ss) + 1
            End If
        Next i

    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then

        For k = 0 To UIDRowPerBlock - 1    ''0...63
            For j = 0 To UIDBitsPerRow - 1 ''0...15, 16 bits per row
                ''''Right-Side
                If DoubleBitArray(k * UIDBitsPerRow + j) <> eFuse_Pgm_Bit(k * UIDReadBitWidth + j) Then
                    FailCnt(ss) = FailCnt(ss) + 1
                End If
                ''''Left-Side
                If DoubleBitArray(k * UIDBitsPerRow + j) <> eFuse_Pgm_Bit(k * UIDReadBitWidth + UIDBitsPerRow + j) Then
                    FailCnt(ss) = FailCnt(ss) + 1
                End If
            Next j
        Next k

    ElseIf (gS_EFuse_Orientation = "SingleUp") Then
 
        For i = 0 To (UIDTotalBits - 1)
            If (DoubleBitArray(i) <> eFuse_Pgm_Bit(i)) Then
                FailCnt(ss) = FailCnt(ss) + 1
            End If
        Next i

    End If
   
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_UIDCompare_DoubleBit_PgmBit_byStage(DoubleBitArray() As Long, eFuse_Pgm_Bit() As Long, FailCnt As SiteLong)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UIDCompare_DoubleBit_PgmBit_byStage"
    
    Dim ss As Variant
    Dim i As Long, j As Long, k As Long
    Dim m_stage As String
    Dim m_startbit As Long
    Dim m_endbit As Long
    Dim bcnt As Long

    ss = TheExec.sites.SiteNumber
    FailCnt(ss) = 0

    For i = 0 To UBound(UIDFuse.Category)
        m_stage = LCase(UIDFuse.Category(i).Stage) ''''<Notice>

        If (gS_JobName = m_stage) Then
            m_startbit = UIDFuse.Category(i).LSBbit
            m_endbit = UIDFuse.Category(i).MSBbit
            
            ''''-------------------------------------------------------------------------------------------
            If (gS_EFuse_Orientation = "SingleUp") Then
                For k = m_startbit To m_endbit
                    If (DoubleBitArray(k) <> eFuse_Pgm_Bit(k)) Then
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
                    If (DoubleBitArray(k) <> eFuse_Pgm_Bit(k + UIDBitsPerBlockUsed)) Then
                        FailCnt(ss) = FailCnt(ss) + 1
                    End If
                Next k
        
            ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
                bcnt = 0 ''''<MUST>
                For k = 0 To UIDRowPerBlock - 1    ''0...63
                    For j = 0 To UIDBitsPerRow - 1 ''0...15, 16 bits per row
                        If (bcnt >= m_startbit And bcnt <= m_endbit) Then
                            ''''Right-Side
                            If DoubleBitArray(k * UIDBitsPerRow + j) <> eFuse_Pgm_Bit(k * UIDReadBitWidth + j) Then
                                FailCnt(ss) = FailCnt(ss) + 1
                            End If
                            ''''Left-Side
                            If DoubleBitArray(k * UIDBitsPerRow + j) <> eFuse_Pgm_Bit(k * UIDReadBitWidth + UIDBitsPerRow + j) Then
                                FailCnt(ss) = FailCnt(ss) + 1
                            End If
                        Else
                            ''''over m_endbit bits, set k,j to up limit to escape for-loop
                            If (bcnt > m_endbit) Then
                                k = UIDRowPerBlock
                                j = UIDBitsPerRow
                            End If
                        End If
                        bcnt = bcnt + 1
                    Next j
                Next k

            End If
            ''''-------------------------------------------------------------------------------------------
        End If ''''end of If (gS_JobName = m_stage)
    Next i

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
''''20150604 New for the new UID eFuse ChkList table
''''Replace the function "AESMakePgmCompareBit()"
Public Function auto_UIDMakePgmCompareBit(ByRef Expanded_eFuse_Pgm_Bit() As Long, ByRef eFuse_Pgm_Bit() As Long, ByRef CompareByte() As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UIDMakePgmCompareBit"

    Dim i As Long, j As Long
    Dim kk As Long, k1 As Long, k2 As Long
    Dim CRCarray(31) As Byte
    Dim jj As Long
    Dim debugout As String
    Dim CompareBit As String, ArrayString As String
    Dim CountN As Long
    Dim ByteString As String
    Dim cnt As Long
    Dim bin_str As String
    Dim eFuse_Read_Bit() As String
    Dim temp_bit() As Long
    
    ReDim eFuse_Read_Bit(UIDTotalBits - 1)
    ReDim temp_bit(UIDTotalBits - 1)

    '=== Initialization ===
    For i = 0 To UBound(eFuse_Pgm_Bit)
        eFuse_Pgm_Bit(i) = 0
    Next i

    '=============================================
    '= Decompose AES0/1/2/3 strings to array     =
    '=============================================
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber        'read current site number
       
    Dim n As Long
    Dim m_stage As String
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant ''''20160616 update, was Long
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim m_bitsum As Long
    Dim m_bitStrM As String
    Dim CRCHex As String
    Dim tmpdlgStr As String
    Dim m_dlgprint As Long
    Dim tmpStr As String

    ''''20150903, new UID category for multiple uid (UID1,UID2_e,FUSE_RAND)
    Dim uidcode_index As Long
    uidcode_index = 0
    
    For i = 0 To UBound(UIDFuse.Category)
        m_stage = LCase(UIDFuse.Category(i).Stage)
        m_catename = UIDFuse.Category(i).Name
        m_algorithm = LCase(UIDFuse.Category(i).algorithm)
        m_LSBbit = UIDFuse.Category(i).LSBbit
        m_MSBBit = UIDFuse.Category(i).MSBbit
        m_bitwidth = UIDFuse.Category(i).BitWidth
        m_defval = UIDFuse.Category(i).DefaultValue
        m_defreal = UIDFuse.Category(i).Default_Real
        m_bitStrM = ""
        m_bitsum = 0
        m_decimal = 0 ''''Must be here for the initilization every time
        m_dlgprint = 0 ''''initialize

        ''''20150710 new datalog format
        tmpdlgStr = "Site(" + CStr(ss) + ") Programming : " + FormatNumeric(m_catename, gI_UID_catename_maxLen)
        tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "(LSB)] = "

        If (m_algorithm = "uid") Then
        
            If (gS_JobName = m_stage) Then
                ''''TMA/Elba: 512/128=4, M8,Cayman: 896/128=7
                ''''for multiple uid categories
                ''''20150903, new UID category for multiple uid (UID1,UID2_e,FUSE_RAND)
                For n = 0 To (m_bitwidth / UIDBitsPerCode) - 1
                    ''''Here m_LSBbit will be increased by 128 (UIDBitsPerCode) inside the below function
                    Call CalcEfuseBitbyString(UID_Code_BitStr(ss, uidcode_index), m_LSBbit, UIDBitsPerCode, eFuse_Pgm_Bit, bin_str)
                    m_bitStrM = m_bitStrM + UID_Code_BitStr(ss, uidcode_index)
                    
                    For j = 1 To UIDBitsPerCode
                        m_bitsum = m_bitsum + CLng(Mid(UID_Code_BitStr(ss, uidcode_index), j, 1))
                    Next j
                    uidcode_index = uidcode_index + 1 ''''<Important>
                Next n
            Else
                m_decimal = 0
                m_bitsum = 0
                
                For j = 1 To m_bitwidth
                    m_bitStrM = "0" + m_bitStrM
                Next j
            End If

            UIDFuse.Category(i).Write.BitStrM(ss) = m_bitStrM
            UIDFuse.Category(i).Write.BitStrL(ss) = StrReverse(m_bitStrM)
            UIDFuse.Category(i).Write.BitSummation(ss) = m_bitsum
            UIDFuse.Category(i).Write.Value(ss) = m_bitsum / m_bitwidth
            UIDFuse.Category(i).Write.ValStr(ss) = CStr(m_bitsum / m_bitwidth)
            
        ElseIf (m_algorithm = "crc") Then
            '=========================================
            '=  Combine UID/AES codes with CRC codes =
            '=========================================
            debugout = ""
            If (gS_JobName = m_stage) Then
            
                Call CRC_Zero_Array(CRCarray) ''''ArW
    
                For j = (gL_UIDCodeBitWidth - 1) To 0 Step -1
                    Call CRC_ComputeCRCforBit(CRCarray, CByte(eFuse_Pgm_Bit(j)))
                    ''debugout = debugout & CStr(eFuse_Pgm_Bit(j))
                Next j
                ''''Debug.Print debugout
                    
                ''''CRC code
                m_bitsum = 0
                m_bitStrM = ""
                For j = 0 To m_bitwidth - 1 ''''here m_bitwidth=32
                    jj = m_LSBbit + j
                    eFuse_Pgm_Bit(jj) = CRCarray(j)
                    If (j < 31) Then m_decimal = m_decimal + (CRCarray(j) * (2 ^ j)) ''''<Notice> j can NOT be 31, otherwise Overflow
                    m_bitsum = m_bitsum + eFuse_Pgm_Bit(jj)
                    m_bitStrM = CStr(eFuse_Pgm_Bit(jj)) + m_bitStrM
                    ''debugout = debugout & CStr(eFuse_Pgm_Bit(jj))
                Next j
                ''''Debug.Print debugout
                CRCHex = ""
                CRCHex = auto_BinStr2HexStr(m_bitStrM, m_bitwidth / 4)

            Else
                m_decimal = 0
                CRCHex = "00000000"
                For j = 0 To m_bitwidth - 1
                    jj = m_LSBbit + j
                    eFuse_Pgm_Bit(jj) = 0
                    m_bitStrM = "0" + m_bitStrM
                Next j
                m_bitsum = 0
            End If

            UIDFuse.Category(i).Write.BitStrM(ss) = m_bitStrM
            UIDFuse.Category(i).Write.BitStrL(ss) = StrReverse(m_bitStrM)
            UIDFuse.Category(i).Write.Decimal(ss) = m_decimal
            UIDFuse.Category(i).Write.BitSummation(ss) = m_bitsum
            UIDFuse.Category(i).Write.Value(ss) = CRCHex
            UIDFuse.Category(i).Write.ValStr(ss) = CRCHex

        ElseIf (m_algorithm = "rid") Then
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("UID", m_catename, m_defreal, m_defval)
            Call auto_eFuse_Dec2PgmArr_Write_byStage("UID", m_stage, i, m_decimal, m_LSBbit, m_MSBBit, eFuse_Pgm_Bit)

        ElseIf (m_algorithm Like "*reserve*") Then
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("UID", m_catename, m_defreal, m_defval)
            Call auto_eFuse_Dec2PgmArr_Write_byStage("UID", m_stage, i, m_decimal, m_LSBbit, m_MSBBit, eFuse_Pgm_Bit)
            m_dlgprint = 1
        ElseIf (m_algorithm = "app") Then
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("UID", m_catename, m_defreal, m_defval)
            Call auto_eFuse_Dec2PgmArr_Write_byStage("UID", m_stage, i, m_decimal, m_LSBbit, m_MSBBit, eFuse_Pgm_Bit)
            m_dlgprint = 1
        Else
            ''''20150720 update
            ''TheExec.Datalog.WriteComment "[WARNING]: unknown UIDFuse Algorithm=" + m_algorithm
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("UID", m_catename, m_defreal, m_defval)
            Call auto_eFuse_Dec2PgmArr_Write_byStage("UID", m_stage, i, m_decimal, m_LSBbit, m_MSBBit, eFuse_Pgm_Bit)
            m_dlgprint = 1
        End If

        If (m_dlgprint = 1) Then
            m_bitStrM = UIDFuse.Category(i).Write.BitStrM(ss)
            m_decimal = UIDFuse.Category(i).Write.Decimal(ss)
            tmpStr = " [" + m_bitStrM + "]"
            tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + tmpStr
            TheExec.Datalog.WriteComment tmpdlgStr
        End If
    Next i

    '========================================================================
    '=      Copy first half 1024 bit data to second half 1024 bit data      =
    '========================================================================
    For i = 0 To UIDTotalBits - 1
        temp_bit(i) = 0
    Next i
        
    If (gS_EFuse_Orientation = "UP2DOWN") Then

        For i = 0 To UIDBitsPerBlockUsed - 1
            eFuse_Pgm_Bit(i + UIDBitsPerBlockUsed) = eFuse_Pgm_Bit(i)
        Next i
        
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then

        For i = 0 To UIDTotalBits - 1 ''''UIDBitsPerBlockUsed - 1
            temp_bit(i) = eFuse_Pgm_Bit(i)
        Next i
        kk = 0 ''''MUST
        ''''k1: Right Side
        ''''k2: Left  Side
        For i = 0 To UIDReadCycle - 1        ''UIDReadCycle=64
            For j = 0 To UIDBitsPerRow - 1   ''UIDBitsPerRow=16
                k1 = (i * UIDReadBitWidth) + j
                k2 = (i * UIDReadBitWidth) + (UIDBitsPerRow + j)
                eFuse_Pgm_Bit(k1) = temp_bit(kk)
                eFuse_Pgm_Bit(k2) = temp_bit(kk)
                kk = kk + 1
            Next j
        Next i

    ElseIf (gS_EFuse_Orientation = "SingleUp") Then
        ''''do nothing here
    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
    End If
    
    '=================================================
    '=  Expand vector number for preparing DSSC Data =
    '=================================================
    'Multiple UIDWriteBitExpandWidth times for DSSC wave
    CountN = 0
    For i = 0 To UBound(eFuse_Pgm_Bit)
        For j = 0 To UIDWriteBitExpandWidth - 1
            Expanded_eFuse_Pgm_Bit(CountN) = eFuse_Pgm_Bit(i)
            CountN = CountN + 1
        Next j
    Next i
  
    '========================================================
    '= Make Data ouptut Q[31:0] cycles for Read pattern     =
    '========================================================
    ''''The below could be unused.
    '============================================
    '=  Copy Pgm Bit array to Read Bit array    =
    '============================================
    For i = 0 To UIDTotalBits - 1
        If eFuse_Pgm_Bit(i) = 0 Then
            eFuse_Read_Bit(i) = "L"
        ElseIf eFuse_Pgm_Bit(i) = 1 Then
            eFuse_Read_Bit(i) = "H"
        Else
            MsgBox ("A fatal problem in VBT " + funcName + ", incorrect eFuse_Pgm_Bit value")
        End If
    Next i
    
    cnt = 0
    For i = 0 To UIDReadCycle - 1
        ByteString = ""
        For j = 0 To UIDReadBitWidth - 1
            ByteString = ByteString + eFuse_Read_Bit(cnt)
            cnt = cnt + 1
        Next j
        CompareByte(i) = StrReverse(ByteString)
    Next i
    ''''-----------------------------------------------------

    auto_UIDMakePgmCompareBit = CountN
        
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20151229 New
Public Function auto_UIDMakePgmCompareBit_byCategory(catename_grp As String, ByRef Expanded_eFuse_Pgm_Bit() As Long, ByRef eFuse_Pgm_Bit() As Long, ByRef CompareByte() As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UIDMakePgmCompareBit_byCategory"

    Dim i As Long, j As Long, n As Long
    Dim kk As Long, k1 As Long, k2 As Long
    Dim CRCarray(31) As Byte
    Dim jj As Long
    Dim debugout As String
    Dim CompareBit As String, ArrayString As String
    Dim CountN As Long
    Dim ByteString As String
    Dim cnt As Long
    Dim bin_str As String
    Dim eFuse_Read_Bit() As String
    Dim temp_bit() As Long
    Dim tmpStr As String
    
    ReDim eFuse_Read_Bit(UIDTotalBits - 1)
    ReDim temp_bit(UIDTotalBits - 1)

    '=== Initialization ===
    For i = 0 To UBound(eFuse_Pgm_Bit)
        eFuse_Pgm_Bit(i) = 0
    Next i

    '=============================================
    '= Decompose AES0/1/2/3 strings to array     =
    '=============================================
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber        'read current site number
       
    Dim m_stage As String
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant ''''20160616 update, was Long
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim m_bitsum As Long
    Dim m_bitStrM As String
    Dim CRCHex As String
    Dim tmpdlgStr As String
    Dim m_dlgprint As Long

    ''''------------------------------------------------
    ''''Split catename_grp as String array
    ''''------------------------------------------------
    Dim m_cateflag As Boolean
    Dim m_catenameArr() As String
    Dim m_cateArr_elem As String
    Dim cateCNT As Long
    m_catenameArr = Split(Trim(catename_grp), ",")
    cateCNT = 0
    For j = 0 To UBound(m_catenameArr)
        m_cateArr_elem = Trim(m_catenameArr(j))
        If (m_cateArr_elem <> "") Then
            m_catenameArr(cateCNT) = m_cateArr_elem
            cateCNT = cateCNT + 1
        End If
    Next j
    If (cateCNT >= 1) Then
        ReDim Preserve m_catenameArr(cateCNT - 1)
    Else
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: There is NO any category assigned, please check it out."
        Exit Function
    End If
    ''''------------------------------------------------

    ''''20150903, new UID category for multiple uid (UID1,UID2_e,FUSE_RAND)
    Dim uidcode_index As Long
    uidcode_index = 0
    
    For i = 0 To UBound(UIDFuse.Category)
        m_stage = LCase(UIDFuse.Category(i).Stage)
        m_catename = UIDFuse.Category(i).Name
        m_algorithm = LCase(UIDFuse.Category(i).algorithm)
        m_LSBbit = UIDFuse.Category(i).LSBbit
        m_MSBBit = UIDFuse.Category(i).MSBbit
        m_bitwidth = UIDFuse.Category(i).BitWidth
        m_defval = UIDFuse.Category(i).DefaultValue
        m_defreal = UIDFuse.Category(i).Default_Real
        m_bitStrM = ""
        m_bitsum = 0
        m_decimal = 0 ''''Must be here for the initilization every time
        m_dlgprint = 0 ''''initialize

        ''''20150710 new datalog format
        tmpdlgStr = "Site(" + CStr(ss) + ") Programming : " + FormatNumeric(m_catename, gI_UID_catename_maxLen)
        tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "(LSB)] = "

        ''''-------------------------------------------------------------------------------
        ''''Check if (m_catename) is the specific Category (catename_grp => m_catenameArr)
        m_cateflag = False
        For j = 0 To UBound(m_catenameArr)
            m_cateArr_elem = UCase(m_catenameArr(j))
            If (m_cateArr_elem = UCase(m_catename)) Then
                m_cateflag = True
                Exit For
            End If
        Next j
        ''''-------------------------------------------------------------------------------

        If (m_algorithm = "uid") Then
        
            If (gS_JobName = m_stage And m_cateflag = True) Then
                ''''TMA/Elba: 512/128=4, M8,Cayman: 896/128=7
                ''''for multiple uid categories
                ''''20150903, new UID category for multiple uid (UID1,UID2_e,FUSE_RAND)
                For n = 0 To (m_bitwidth / UIDBitsPerCode) - 1
                    ''''Here m_LSBbit will be increased by 128 (UIDBitsPerCode) inside the below function
                    Call CalcEfuseBitbyString(UID_Code_BitStr(ss, uidcode_index), m_LSBbit, UIDBitsPerCode, eFuse_Pgm_Bit, bin_str)
                    m_bitStrM = m_bitStrM + UID_Code_BitStr(ss, uidcode_index)
                    
                    For j = 1 To UIDBitsPerCode
                        m_bitsum = m_bitsum + CLng(Mid(UID_Code_BitStr(ss, uidcode_index), j, 1))
                    Next j
                    uidcode_index = uidcode_index + 1 ''''<Important>
                Next n
            Else
                m_decimal = 0
                m_bitsum = 0
                
                For j = 1 To m_bitwidth
                    m_bitStrM = "0" + m_bitStrM
                Next j
            End If

            UIDFuse.Category(i).Write.BitStrM(ss) = m_bitStrM
            UIDFuse.Category(i).Write.BitStrL(ss) = StrReverse(m_bitStrM)
            UIDFuse.Category(i).Write.BitSummation(ss) = m_bitsum
            UIDFuse.Category(i).Write.Value(ss) = m_bitsum / m_bitwidth
            UIDFuse.Category(i).Write.ValStr(ss) = CStr(m_bitsum / m_bitwidth)
            
        ElseIf (m_algorithm = "crc") Then
            '=========================================
            '=  Combine UID/AES codes with CRC codes =
            '=========================================
            
            debugout = ""
            If (gS_JobName = m_stage And m_cateflag = True) Then
            
                Call CRC_Zero_Array(CRCarray) ''''ArW
    
                For j = (gL_UIDCodeBitWidth - 1) To 0 Step -1
                    Call CRC_ComputeCRCforBit(CRCarray, CByte(eFuse_Pgm_Bit(j)))
                    ''debugout = debugout & CStr(eFuse_Pgm_Bit(j))
                Next j
                ''''Debug.Print debugout
                    
                ''''CRC code
                m_bitsum = 0
                m_bitStrM = ""
                For j = 0 To m_bitwidth - 1 ''''here m_bitwidth=32
                    jj = m_LSBbit + j
                    eFuse_Pgm_Bit(jj) = CRCarray(j)
                    If (j < 31) Then m_decimal = m_decimal + (CRCarray(j) * (2 ^ j)) ''''<Notice> j can NOT be 31, otherwise Overflow
                    m_bitsum = m_bitsum + eFuse_Pgm_Bit(jj)
                    m_bitStrM = CStr(eFuse_Pgm_Bit(jj)) + m_bitStrM
                    ''debugout = debugout & CStr(eFuse_Pgm_Bit(jj))
                Next j
                ''''Debug.Print debugout
                CRCHex = ""
                CRCHex = auto_BinStr2HexStr(m_bitStrM, m_bitwidth / 4)

            Else
                m_decimal = 0
                CRCHex = "00000000"
                For j = 0 To m_bitwidth - 1
                    jj = m_LSBbit + j
                    eFuse_Pgm_Bit(jj) = 0
                    m_bitStrM = "0" + m_bitStrM
                Next j
                m_bitsum = 0
            End If

            UIDFuse.Category(i).Write.BitStrM(ss) = m_bitStrM
            UIDFuse.Category(i).Write.BitStrL(ss) = StrReverse(m_bitStrM)
            UIDFuse.Category(i).Write.Decimal(ss) = m_decimal
            UIDFuse.Category(i).Write.BitSummation(ss) = m_bitsum
            UIDFuse.Category(i).Write.Value(ss) = CRCHex
            UIDFuse.Category(i).Write.ValStr(ss) = CRCHex

        ElseIf (m_algorithm = "rid") Then
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("UID", m_catename, m_defreal, m_defval)
            Call auto_eFuse_Dec2PgmArr_Write_byCategory("UID", m_cateflag, i, m_decimal, m_LSBbit, m_MSBBit, eFuse_Pgm_Bit)
            
        ElseIf (m_algorithm Like "*reserve*") Then
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("UID", m_catename, m_defreal, m_defval)
            Call auto_eFuse_Dec2PgmArr_Write_byCategory("UID", m_cateflag, i, m_decimal, m_LSBbit, m_MSBBit, eFuse_Pgm_Bit)
            m_dlgprint = 1
        ElseIf (m_algorithm = "app") Then
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("UID", m_catename, m_defreal, m_defval)
            Call auto_eFuse_Dec2PgmArr_Write_byCategory("UID", m_cateflag, i, m_decimal, m_LSBbit, m_MSBBit, eFuse_Pgm_Bit)
            m_dlgprint = 1
        Else
            ''''20150720 update
            ''TheExec.Datalog.WriteComment "[WARNING]: unknown UIDFuse Algorithm=" + m_algorithm
            m_decimal = auto_eFuse_Get_DefaultRealDecimal("UID", m_catename, m_defreal, m_defval)
            Call auto_eFuse_Dec2PgmArr_Write_byCategory("UID", m_cateflag, i, m_decimal, m_LSBbit, m_MSBBit, eFuse_Pgm_Bit)
            m_dlgprint = 1
        End If

        If (m_dlgprint = 1) Then
            m_bitStrM = UIDFuse.Category(i).Write.BitStrM(ss)
            m_decimal = UIDFuse.Category(i).Write.Decimal(ss)
            tmpStr = " [" + m_bitStrM + "]"
            tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + tmpStr
            TheExec.Datalog.WriteComment tmpdlgStr
        End If
    Next i
    
    '========================================================================
    '=      Copy first half 1024 bit data to second half 1024 bit data      =
    '========================================================================
    For i = 0 To UIDTotalBits - 1
        temp_bit(i) = 0
    Next i
        
    If (gS_EFuse_Orientation = "UP2DOWN") Then

        For i = 0 To UIDBitsPerBlockUsed - 1
            eFuse_Pgm_Bit(i + UIDBitsPerBlockUsed) = eFuse_Pgm_Bit(i)
        Next i
        
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then

        For i = 0 To UIDTotalBits - 1 ''''UIDBitsPerBlockUsed - 1
            temp_bit(i) = eFuse_Pgm_Bit(i)
        Next i
        kk = 0 ''''MUST
        ''''k1: Right Side
        ''''k2: Left  Side
        For i = 0 To UIDReadCycle - 1        ''UIDReadCycle=64
            For j = 0 To UIDBitsPerRow - 1   ''UIDBitsPerRow=16
                k1 = (i * UIDReadBitWidth) + j
                k2 = (i * UIDReadBitWidth) + (UIDBitsPerRow + j)
                eFuse_Pgm_Bit(k1) = temp_bit(kk)
                eFuse_Pgm_Bit(k2) = temp_bit(kk)
                kk = kk + 1
            Next j
        Next i

    ElseIf (gS_EFuse_Orientation = "SingleUp") Then
        ''''do nothing here
    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
    End If
    
    '=================================================
    '=  Expand vector number for preparing DSSC Data =
    '=================================================
    'Multiple UIDWriteBitExpandWidth times for DSSC wave
    CountN = 0
    For i = 0 To UBound(eFuse_Pgm_Bit)
        For j = 0 To UIDWriteBitExpandWidth - 1
            Expanded_eFuse_Pgm_Bit(CountN) = eFuse_Pgm_Bit(i)
            CountN = CountN + 1
        Next j
    Next i
  
    '========================================================
    '= Make Data ouptut Q[31:0] cycles for Read pattern     =
    '========================================================
    ''''The below could be unused.
    '============================================
    '=  Copy Pgm Bit array to Read Bit array    =
    '============================================
    For i = 0 To UIDTotalBits - 1
        If eFuse_Pgm_Bit(i) = 0 Then
            eFuse_Read_Bit(i) = "L"
        ElseIf eFuse_Pgm_Bit(i) = 1 Then
            eFuse_Read_Bit(i) = "H"
        Else
            MsgBox ("A fatal problem in VBT " + funcName + ", incorrect eFuse_Pgm_Bit value")
        End If
    Next i
    
    cnt = 0
    For i = 0 To UIDReadCycle - 1
        ByteString = ""
        For j = 0 To UIDReadBitWidth - 1
            ByteString = ByteString + eFuse_Read_Bit(cnt)
            cnt = cnt + 1
        Next j
        CompareByte(i) = StrReverse(ByteString)
    Next i
    ''''-----------------------------------------------------

    auto_UIDMakePgmCompareBit_byCategory = CountN
        
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function CalcEfuseBitbyString(DataString As String, ByRef start_bit As Long, length As Long, ByRef eFuse_Pgm_Bit() As Long, ByRef bin_str As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CalcEfuseBitbyString"
    
    Dim StringLen As Long
    
    Dim j As Long
    
    StringLen = Len(DataString)
    If length <> StringLen Then
    TheExec.Datalog.WriteComment "Data length not match"
    End If
    bin_str = ""
    '*** Most Left bit of AES/UID string is MSB ***
'    For j = StringLen - 1 To 0 Step -1
'        eFuse_Pgm_Bit(start_bit) = CLng(Mid(DataString, j + 1, 1)) 'BinArray(j)
'        bin_str = bin_str + Mid(DataString, j + 1, 1)
'        start_bit = start_bit + 1
'    Next j
    
    '*** Most Left bit of AES/UID string is LSB (C651's Email on 6/3) ***
    For j = 0 To StringLen - 1
        eFuse_Pgm_Bit(start_bit) = CLng(Mid(DataString, j + 1, 1)) 'BinArray(j)
        bin_str = bin_str + Mid(DataString, j + 1, 1)
        start_bit = start_bit + 1
    Next j

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_Decode_UIDBinary_Data(SrcArray() As Long, Optional showPrint As Boolean = True)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Decode_UIDBinary_Data"

    Dim PartialStr As String
    Dim i As Long, j As Long
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber
        
    Dim m_bitStrM As String
    Dim m_catename As String
    Dim m_resolution As Double
    Dim m_algorithm As String
    Dim m_MSBBit As Long
    Dim m_LSBbit As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant ''''20160506 update, was Long
    Dim m_defreal As String
    Dim OutputStr As String
    Dim m_defval As Variant
    Dim m_bitsum As Long
    Dim CRCHex As String
    Dim TmpVal As Variant
    Dim tmpStr As String
    Dim m_HexStr As String
    
    For i = 0 To UBound(UIDFuse.Category)
        PartialStr = ""
        OutputStr = ""
        m_bitStrM = ""
        m_decimal = 0
        CRCHex = ""

        m_catename = UIDFuse.Category(i).Name
        m_algorithm = LCase(UIDFuse.Category(i).algorithm)
        m_LSBbit = UIDFuse.Category(i).LSBbit
        m_MSBBit = UIDFuse.Category(i).MSBbit
        m_bitwidth = UIDFuse.Category(i).BitWidth
        m_defval = UIDFuse.Category(i).DefaultValue
        m_defreal = UIDFuse.Category(i).Default_Real

        ''''20150825 update
        Call auto_eFuse_Bin2DecStr("UID", i, SrcArray, m_LSBbit, m_MSBBit)
        
        m_bitStrM = UIDFuse.Category(i).Read.BitStrM(ss)
        m_decimal = UIDFuse.Category(i).Read.Decimal(ss)
        m_bitsum = UIDFuse.Category(i).Read.BitSummation(ss)
        
        tmpStr = " [" + m_bitStrM + "]"
        PartialStr = FormatNumeric(m_catename, gI_UID_catename_maxLen) + " [(MSB)" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "(LSB)] = "

        If (m_algorithm = "uid") Then
            TmpVal = m_bitsum / m_bitwidth
            ''''TmpStr = Mid(CStr(tmpVal), 1, 8)
            tmpStr = Format(TmpVal, "0.000000") ''''20151229 update
            UIDFuse.Category(i).Read.Value(ss) = TmpVal
            UIDFuse.Category(i).Read.ValStr(ss) = tmpStr
            
            PartialStr = PartialStr + FormatNumeric(tmpStr, 10)
            
        ElseIf (m_algorithm = "crc") Then
            ''''the below has been included inside auto_eFuse_Bin2DecStr()
            ''CRCHex = auto_BinStr2HexStr(m_bitstrM, m_bitwidth / 4)
            ''UIDFuse.Category(i).Read.Value(ss) = CRCHex
            ''UIDFuse.Category(i).Read.ValStr(ss) = CRCHex
            ''UIDFuse.Category(i).Read.HexStr(ss) = "0x" + CRCHex ''''20161013
            PartialStr = PartialStr + FormatNumeric(UIDFuse.Category(i).Read.HexStr(ss), 10)
            
''''        ElseIf (m_algorithm = "rid") Then
''''            PartialStr = PartialStr + FormatNumeric(m_decimal, 10)
''''
''''        ElseIf (m_algorithm Like "*reserve*") Then
''''            PartialStr = PartialStr + FormatNumeric(m_decimal, 10) ''''20160927 update, was m_bitsum
''''
''''        ElseIf (m_algorithm = "app") Then
''''            PartialStr = PartialStr + FormatNumeric(m_decimal, 10)

        Else
            m_HexStr = UIDFuse.Category(i).Read.HexStr(ss) ''''here m_hexstr with prefix "0x"
            tmpStr = FormatNumeric(tmpStr, -35) + FormatNumeric(" [" + m_HexStr + "]", -10)
            ''''undefined Algorithm
            ''''other cases, 20160927 update
            PartialStr = PartialStr + FormatNumeric(m_decimal, 10) + tmpStr ''''+ " (undefined Algorithm " + m_algorithm + ")"
        End If
        
        OutputStr = "Site(" + CStr(ss) + ") Read from DSSC : "
        If (showPrint) Then TheExec.Datalog.WriteComment OutputStr + PartialStr
        
    Next i ''''for loop UBound(UIDFuse.Category)
        
    TheExec.Datalog.WriteComment ""
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
  
End Function
