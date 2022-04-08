Attribute VB_Name = "DSP_LIB_eFuse"
Option Explicit

Public Function eFuse_DspWave_Copy(ByVal InWave As DSPWave, ByRef outWave As DSPWave) As Long
On Error Resume Next

    outWave = InWave.ConvertDataTypeTo(DspLong).Copy
    
End Function

Private Function eFuse_reverseBitWave(ByVal InWave As DSPWave, ByRef outWave As DSPWave) As Long
On Error Resume Next
    
    Dim i As Long
    Dim m_size As Long
    Dim m_tmpArr() As Long
    Dim m_outArr() As Long
    Dim m_tmpWave As New DSPWave

    m_tmpWave = InWave.Copy.ConvertDataTypeTo(DspLong)

    m_size = InWave.SampleSize
    outWave.CreateConstant 0, m_size, DspLong
    m_tmpArr = m_tmpWave.Data
    m_outArr = outWave.Data
    For i = 0 To m_size - 1
        ''outWave.Element(i) = m_tmpWave.Element(m_size - i - 1) ''''waste TT
        m_outArr(i) = m_tmpArr(m_size - i - 1) ''''save TT
    Next i
    
    outWave.Data = m_outArr ''''save TT
    
End Function

Private Function eFuse_extract32bits_perValue(ByVal j_32bits As Long, ByVal inValue As Double, ByVal BitWidth As Long, ByRef outWave As DSPWave, outValue As Double) As Long
On Error Resume Next
    
    Dim k_dbl As Double
    Dim m_decimal As Double
    Dim m_32bits_base As Double
    ''''quo:quotient, mod:remainder
    Dim m_dbl_quo As Double
    Dim m_dbl_mod As Double
    Dim m_dummy As Double
    Dim m_parWave As New DSPWave
    
    m_32bits_base = CDbl(2 ^ 32)
    
    m_decimal = inValue
    k_dbl = CDbl(m_32bits_base ^ j_32bits)
    
    m_dbl_quo = m_decimal / k_dbl
    m_dummy = Floor(m_dbl_quo) * k_dbl
    ''m_dbl_mod = CDbl(m_decimal - (Floor(m_dbl_quo) * k_dbl))
    m_dbl_mod = CDbl(m_decimal - m_dummy)
    Debug.Print j_32bits & ", " & m_decimal & " / " & k_dbl & " = (quo) " & m_dbl_quo & " ...(mod) " & m_dbl_mod & " , " & (m_dbl_quo * k_dbl)
    
    m_parWave.CreateConstant 0, 1, DspDouble
    
    m_parWave.Element(0) = m_dbl_quo
    outWave = m_parWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong).Copy
    
    outValue = m_dbl_mod
    
End Function

''''can NOT handle/process numOfbits over 52bits well, be used carefully
Private Function eFuse_Value2BinaryWave(ByVal inValue As Double, ByVal BitWidth As Long, ByRef outWave As DSPWave, revBin As Boolean) As Long
On Error Resume Next
    ''''------------------------------------------------------------------------------
    ''''Definition of the Variables
    ''''------------------------------------------------------------------------------
    ''''inValue : input double value
    ''''bitwidth: bin width of the output binary wave
    ''''outWave : the output binary wave
    ''''------------------------------------------------------------------------------
    
    Dim i As Long, j As Long, k As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Double
    
    ''''quo:quotient, mod:remainder
    Dim m_dbl_quo As Double
    Dim m_dbl_mod As Double

    Dim m_hex_quo As Long
    Dim m_hex_mod As Long
    Dim m_32bits_quo As Long
    Dim m_32bits_mod As Long
    
    Dim m_parWave As New DSPWave
    Dim m_par2serWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    
    m_decimal = inValue
    m_bitwidth = BitWidth
    m_hex_quo = Floor(m_bitwidth / 4)
    m_hex_mod = m_bitwidth Mod 4

    m_parWave.CreateConstant 0, 1, DspLong
    m_par2serWave.CreateConstant 0, m_bitwidth, DspLong
    m_tmpWave.CreateConstant 0, m_bitwidth, DspLong
    outWave.CreateConstant 0, m_bitwidth, DspLong
    
    If (m_bitwidth <= 32) Then
        ''''get LSB first for the input number
        For j = 0 To m_hex_quo - 1
            ''''quotient and mod
            m_dbl_quo = Floor(m_decimal / 16)
            m_dbl_mod = m_decimal - (m_dbl_quo * 16#)
            Debug.Print j & ", " & m_decimal & ", " & m_dbl_quo & ", " & m_dbl_mod
            ''''here m_dbl_mod is the value 0~15
            m_parWave.Element(0) = m_dbl_mod
            m_par2serWave = m_parWave.ConvertStreamTo(tldspSerial, 4, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong).Copy
            m_decimal = m_dbl_quo
            Call m_tmpWave.Select(j * 4, 1, 4).Replace(m_par2serWave)
        Next j
    
        If (m_hex_mod > 0) Then
            k = m_hex_quo * 4
            m_parWave.Element(0) = m_dbl_quo
            m_par2serWave = m_parWave.ConvertStreamTo(tldspSerial, m_hex_mod, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong).Copy
            Call m_tmpWave.Select(k, 1, m_hex_mod).Replace(m_par2serWave)
        End If
    
    Else '''' > 32bits
        Dim k_dbl As Double
        Dim m_32bits_base As Double
        Dim m_tmpWave32 As New DSPWave
        Dim m_decWave As New DSPWave
        
        m_32bits_quo = Floor(m_bitwidth / 32)
        m_32bits_mod = m_bitwidth Mod 32

        If (m_32bits_mod = 0) Then
            m_tmpWave32.CreateConstant 0, m_bitwidth, DspLong
            m_32bits_quo = m_32bits_quo - 1
        Else
            m_tmpWave32.CreateConstant 0, (m_32bits_quo + 1) * 32, DspLong
        End If

        ''m_32bits_base = CDbl(2 ^ 32) ''''=4294967296
        
        Dim m_nextValue As Double
        
        ''''get MSB first for the larger number
        For j = m_32bits_quo To 1 Step -1
            
            Call eFuse_extract32bits_perValue(j, m_decimal, 32, m_par2serWave, m_nextValue)
            
            m_decimal = m_nextValue
            Debug.Print m_decimal & " , " & m_nextValue
'''            ''''quotient and mod
'''            k_dbl = CDbl(m_32bits_base ^ j)
'''            m_dbl_quo = Floor(m_decimal / k_dbl)
'''            m_dbl_mod = m_decimal - CDbl(m_dbl_quo * k_dbl)
'''            Debug.Print j & ", " & m_decimal & " / " & k_dbl & " = (quo) " & m_dbl_quo & " ...(mod) " & m_dbl_mod
'''            m_parWave.CreateConstant 0, 1, DspDouble
'''            m_parWave.Element(0) = m_dbl_quo
'''            m_par2serWave = m_parWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong).Copy
'''            m_decimal = m_dbl_mod
            
            Call m_tmpWave32.Select(j * 32, 1, 32).Replace(m_par2serWave)
        Next j
        
        If (m_32bits_mod > 0) Then
            m_parWave.CreateConstant 0, 1, DspDouble
            m_parWave.Element(0) = m_dbl_mod
            m_par2serWave = m_parWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong).Copy
            Call m_tmpWave32.Select(0, 1, 32).Replace(m_par2serWave)
        End If
        m_tmpWave = m_tmpWave32.Select(0, 1, m_bitwidth).Copy
    End If

    If (revBin = True) Then
        Call eFuse_reverseBitWave(m_tmpWave, outWave)
    Else
        outWave = m_tmpWave.Copy
    End If
    
End Function

Private Function eFuse_compWave(ByVal cmp1Wave As DSPWave, ByVal cmp2Wave As DSPWave, ByRef outWave As DSPWave, result As Boolean) As Long
On Error Resume Next

    Dim m_size As Long
    Dim m_sum As Long
    
    m_size = cmp1Wave.SampleSize
    outWave = cmp1Wave.LogicalCompare(EqualTo, cmp2Wave).Copy
    m_sum = outWave.CalcSum
    If (m_sum = m_size) Then
        result = True  ''''Pass
    Else
        result = False ''''Fail
    End If
    
End Function

Private Function eFuse_Gen_DoubleBitWave(ByVal singleBitWave As DSPWave, ByRef doubleBitWave As DSPWave, _
                                         ByRef singleBitSum As Long, ByRef doubleBitSum As Long, ByRef fbcSum As Long, _
                                         ByRef cmpsgWaveperCyc As DSPWave) As Long
On Error Resume Next

    Dim i As Long, j As Long, k As Long
    Dim k1 As Long, k2 As Long
    Dim kk As Long
    
    Dim BitsPerRow As Long
    Dim ReadCycles As Long
    Dim BitsPerCycle As Long
    Dim BitsPerBlock As Long
    
    Dim m_tmpWave1 As New DSPWave
    Dim m_tmpWave2 As New DSPWave

    Dim m_cmpsgMLWave As New DSPWave        ''''it's a bit2bit comparison
    Dim m_cmpsgResultWave As New DSPWave    ''''it's a result of 16bits comparison per cycle
    Dim m_sgbitArrL() As Long ''''LSB side if 2-bit mode (bit15...0 )
    Dim m_sgbitArrM() As Long ''''MSB side if 2-bit mode (bit31...16)

    Dim m_dbbitArr() As Long
    Dim m_sgBitArr() As Long
    
    BitsPerRow = gDL_BitsPerRow
    ReadCycles = gDL_ReadCycles
    BitsPerCycle = gDL_BitsPerCycle
    BitsPerBlock = gDL_BitsPerBlock

    ''''<Important>
    doubleBitWave.CreateConstant 0, BitsPerBlock, DspLong
    doubleBitSum = 0 ''''MUST
    
    If (gDL_eFuse_Orientation = 0) Then ''''Up2Down
        m_tmpWave1 = singleBitWave.Select(0, 1, BitsPerBlock).Copy
        m_tmpWave2 = singleBitWave.Select(BitsPerBlock, 1, BitsPerBlock).Copy
        doubleBitWave = m_tmpWave1.BitwiseOr(m_tmpWave2)

    ElseIf (gDL_eFuse_Orientation = 1) Then ''''1-Bit, SingleUp
        ''''doubleBitWave is equal to singleBitWave
        doubleBitWave = singleBitWave.Copy
        
    ElseIf (gDL_eFuse_Orientation = 2) Then ''''2-Bit, Right2Left
        m_sgBitArr = singleBitWave.Data
        m_dbbitArr = doubleBitWave.Data
        m_sgbitArrL = doubleBitWave.Data
        m_sgbitArrM = doubleBitWave.Data
        
        k = 0 ''''must be here
        For i = 0 To ReadCycles - 1      ''0...15(ECID), 0...31(CFG,SEN)
            For j = 0 To BitsPerRow - 1  ''0...15(ECID), 0...15(CFG,SEN)
                ''''k1: Right block (LSB side, bit15...0 )
                ''''k2:  Left block (MSB side, bit31...16)
                k1 = (i * BitsPerCycle) + j ''<Important> Must use BitsPerCycle here
                k2 = (i * BitsPerCycle) + BitsPerRow + j
                ''doubleBitWave.Element(k) = singleBitWave.Element(k1) Or singleBitWave.Element(k2) ''''waste TT
                m_dbbitArr(k) = m_sgBitArr(k1) Or m_sgBitArr(k2) ''''save TT
                m_sgbitArrL(k) = m_sgBitArr(k1)
                m_sgbitArrM(k) = m_sgBitArr(k2)
                k = k + 1
            Next j
        Next i
        doubleBitWave.Data = m_dbbitArr ''''save TT
        m_tmpWave1.Data = m_sgbitArrL
        m_tmpWave2.Data = m_sgbitArrM

    ''''Below is reserved for the future, there is NO any definition at present.
    ElseIf (gDL_eFuse_Orientation = 3) Then
    ElseIf (gDL_eFuse_Orientation = 4) Then
    ElseIf (gDL_eFuse_Orientation = 5) Then
    End If

    ''''calculate the Sum
    singleBitSum = singleBitWave.CalcSum
    doubleBitSum = doubleBitWave.CalcSum
    If (gDL_eFuse_Orientation = 1) Then ''''1-Bit, SingleUp
        fbcSum = singleBitSum - doubleBitSum
    Else
        fbcSum = singleBitSum - 2 * doubleBitSum
    End If

    ''''201811XX to meet HDC team request as MC2T project
    m_cmpsgResultWave.CreateConstant 1, ReadCycles, DspLong

    ''''2-bit mode, used to check if both MSB side and LSB side are equal
    If (gDL_eFuse_Orientation = 2) Then
        m_cmpsgMLWave = m_tmpWave1.LogicalCompare(EqualTo, m_tmpWave2)
        If (gDB_SerialType = False) Then ''20190801
            m_cmpsgResultWave = m_cmpsgMLWave.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsMsb)
            m_cmpsgResultWave = m_cmpsgResultWave.Divide(65535).Floor(1) ''''(2^16 - 1)=65535
        End If
    End If
    cmpsgWaveperCyc = m_cmpsgResultWave.Copy

End Function

'Private Function eFuse_decode_DSSCReadWave(ByVal FuseType As Long, ByVal inWave As DSPWave, ByRef outWave As DSPWave) As Long
Public Function eFuse_decode_DSSCReadWave(ByVal FuseType As Long, ByVal InWave As DSPWave, ByRef outWave As DSPWave) As Long
On Error Resume Next
     
    ''''------------------------------------------------------------------------------
    ''''Definition of the Variables
    ''''------------------------------------------------------------------------------
    ''''inWave:   eFuse DSSC DoubleBitWave
    ''''------------------------------------------------------------------------------

    Dim i As Long, j As Long
    Dim m_size As Long
    Dim m_MSBBit As Long
    Dim m_LSBbit As Long
    Dim m_bitwidth As Long
    Dim m_defreal As Long
    Dim m_decimal As Double
    Dim m_bitsum As Long

    Dim m_tmpWave As New DSPWave
    Dim m_msbCateWave As New DSPWave
    Dim m_lsbCateWave As New DSPWave
    Dim m_bitwidthCateWave As New DSPWave
    Dim m_defrealCateWave As New DSPWave

    Dim outArr() As Double
    Dim m_tmpArr() As Long
    Dim m_msbCateArr() As Long
    Dim m_lsbCateArr() As Long
    Dim m_bitwidthCateArr() As Long
    Dim m_defrealCateArr() As Long

''''''''-----------------------------------
''''''''From Enum eFuseBlockType
''''    eFuse_ECID = 1
''''    eFuse_CFG = 2
''''    eFuse_UID = 3
''''    eFuse_SEN = 4
''''    eFuse_MON = 5
''''    eFuse_UDR = 6
''''    eFuse_UDRE = 7
''''    eFuse_UDRP = 8
''''    eFuse_CMP = 9
''''    eFuse_CMPE = 10
''''    eFuse_CMPP = 11
''''    eFuse_Block_Unknown = 999
''''''''-----------------------------------

    If (FuseType = 1) Then
        m_msbCateArr = gDW_ECID_MSBBit_Cate.Data
        m_lsbCateArr = gDW_ECID_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_ECID_BitWidth_Cate.Data
        m_defrealCateArr = gDW_ECID_DefaultReal_Cate.Data

    ElseIf (FuseType = 2) Then
        m_msbCateArr = gDW_CFG_MSBBit_Cate.Data
        m_lsbCateArr = gDW_CFG_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_CFG_BitWidth_Cate.Data
        m_defrealCateArr = gDW_CFG_DefaultReal_Cate.Data

    ElseIf (FuseType = 3) Then
        m_msbCateArr = gDW_UID_MSBBit_Cate.Data
        m_lsbCateArr = gDW_UID_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_UID_BitWidth_Cate.Data
        m_defrealCateArr = gDW_UID_DefaultReal_Cate.Data
    ElseIf (FuseType = 4) Then
        m_msbCateArr = gDW_SEN_MSBBit_Cate.Data
        m_lsbCateArr = gDW_SEN_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_SEN_BitWidth_Cate.Data
        m_defrealCateArr = gDW_SEN_DefaultReal_Cate.Data
    ElseIf (FuseType = 5) Then
        m_msbCateArr = gDW_MON_MSBBit_Cate.Data
        m_lsbCateArr = gDW_MON_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_MON_BitWidth_Cate.Data
        m_defrealCateArr = gDW_MON_DefaultReal_Cate.Data
    ElseIf (FuseType = 6) Then
        m_msbCateArr = gDW_UDR_MSBBit_Cate.Data
        m_lsbCateArr = gDW_UDR_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_UDR_BitWidth_Cate.Data
        m_defrealCateArr = gDW_UDR_DefaultReal_Cate.Data
    ElseIf (FuseType = 7) Then
        m_msbCateArr = gDW_UDRE_MSBBit_Cate.Data
        m_lsbCateArr = gDW_UDRE_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_UDRE_BitWidth_Cate.Data
        m_defrealCateArr = gDW_UDRE_DefaultReal_Cate.Data
    ElseIf (FuseType = 8) Then
        m_msbCateArr = gDW_UDRP_MSBBit_Cate.Data
        m_lsbCateArr = gDW_UDRP_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_UDRP_BitWidth_Cate.Data
        m_defrealCateArr = gDW_UDRP_DefaultReal_Cate.Data
    ElseIf (FuseType = 9) Then
        m_msbCateArr = gDW_CMP_MSBBit_Cate.Data
        m_lsbCateArr = gDW_CMP_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_CMP_BitWidth_Cate.Data
        m_defrealCateArr = gDW_CMP_DefaultReal_Cate.Data
    ElseIf (FuseType = 10) Then
        m_msbCateArr = gDW_CMPE_MSBBit_Cate.Data
        m_lsbCateArr = gDW_CMPE_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_CMPE_BitWidth_Cate.Data
        m_defrealCateArr = gDW_CMPE_DefaultReal_Cate.Data
    ElseIf (FuseType = 11) Then
        m_msbCateArr = gDW_CMPP_MSBBit_Cate.Data
        m_lsbCateArr = gDW_CMPP_LSBBit_Cate.Data
        m_bitwidthCateArr = gDW_CMPP_BitWidth_Cate.Data
        m_defrealCateArr = gDW_CMPP_DefaultReal_Cate.Data
    Else
    End If
    
    m_size = UBound(m_msbCateArr) + 1 ''''m_msbCateWave.SampleSize
    outWave.CreateConstant 0, m_size, DspDouble
    outArr = outWave.Data

    If (FuseType = 1) Then
        ''''ECID: bit location msb < lsb, so use "Bit0IsLsb" convert to the decimal
        'Dim m_user_proberSite As New SiteVariant
        For i = 0 To m_size - 1
            m_MSBBit = m_msbCateArr(i)
            m_LSBbit = m_lsbCateArr(i)
            m_bitwidth = m_bitwidthCateArr(i)
            m_tmpWave.CreateConstant 0, m_bitwidth, DspLong
            m_tmpWave = InWave.Select(m_MSBBit, 1, m_bitwidth).Copy
            m_bitsum = m_tmpWave.CalcSum
            
            m_decimal = m_tmpWave.ConvertStreamTo(tldspParallel, m_bitwidth, 0, Bit0IsLsb).ConvertDataTypeTo(DspDouble).Element(0)
            'm_user_proberSite = auto_WaferData_to_HexECID_SiteAware(64)
            
'            If (m_bitwidth <= 32 Or m_bitSum = 0) Then
'                m_decimal = m_tmpWave.ConvertStreamTo(tldspParallel, m_bitwidth, 0, Bit0IsLsb).ConvertDataTypeTo(DspDouble).Element(0)
'            Else
'                ''''Here it means that bitwidth is > 32bits and NOT zero.
'                ''''it will be present by Hex and Binary compare with the limit.
'                m_decimal = -9999
'            End If
            outArr(i) = m_decimal

            ''''trial
            ''Call eFuse_Value2BinaryWave(m_decimal, m_bitwidth, out2Wave, True) ''''ECID need to reverse Binary [LSB...MSB]
            ''Call eFuse_compWave(out2Wave, m_tmpWave2, m_cmpWave, m_Result)
        Next i
    Else
        ''''Others Fuse: bit location msb > lsb, so use "Bit0IsMsb" convert to the decimal
        For i = 0 To m_size - 1
            m_MSBBit = m_msbCateArr(i)
            m_LSBbit = m_lsbCateArr(i)
            m_bitwidth = m_bitwidthCateArr(i)
            m_tmpWave.CreateConstant 0, m_bitwidth, DspLong
            m_tmpWave = InWave.Select(m_LSBbit, 1, m_bitwidth).Copy
            m_bitsum = m_tmpWave.CalcSum
            
            If (m_bitwidth <= 32 Or m_bitsum = 0) Then
                m_decimal = m_tmpWave.ConvertStreamTo(tldspParallel, m_bitwidth, 0, Bit0IsMsb).ConvertDataTypeTo(DspDouble).Element(0)
            Else
                ''''Here it means that bitwidth is > 32bits and NOT zero.
                ''''it will be present by Hex and Binary compare with the limit.
                m_decimal = -9999
            End If
            outArr(i) = m_decimal
            
            ''Call eFuse_Value2BinaryWave(m_decimal, m_bitwidth, out2Wave, False) ''''Others Binary [MSB...LSB]
            ''Call eFuse_compWave(out2Wave, m_tmpWave2, m_cmpWave, m_Result)
        Next i
    End If
    outWave.Data = outArr

    ''''update to global DSP variable
    If (FuseType = 1) Then
        gDW_ECID_Read_Decimal_Cate = outWave.Copy
    ElseIf (FuseType = 2) Then
        gDW_CFG_Read_Decimal_Cate = outWave.Copy
    ElseIf (FuseType = 3) Then
        gDW_UID_Read_Decimal_Cate = outWave.Copy
    ElseIf (FuseType = 4) Then
        gDW_SEN_Read_Decimal_Cate = outWave.Copy
    ElseIf (FuseType = 5) Then
        gDW_MON_Read_Decimal_Cate = outWave.Copy
    ElseIf (FuseType = 6) Then
        gDW_UDR_Read_Decimal_Cate = outWave.Copy
    ElseIf (FuseType = 7) Then
        gDW_UDRE_Read_Decimal_Cate = outWave.Copy
    ElseIf (FuseType = 8) Then
        gDW_UDRP_Read_Decimal_Cate = outWave.Copy
    ElseIf (FuseType = 9) Then
        gDW_CMP_Read_Decimal_Cate = outWave.Copy
    ElseIf (FuseType = 10) Then
        gDW_CMPE_Read_Decimal_Cate = outWave.Copy
    ElseIf (FuseType = 11) Then
        gDW_CMPP_Read_Decimal_Cate = outWave.Copy
    Else
    End If
    
End Function

''''201812XX update
Public Function eFuse_Wave32bits_to_SingleDoubleBitWave(ByVal FuseType As Long, ByVal bitFlag_mode As Long, ByVal InWave As DSPWave, _
                                                        ByRef fbcSum As Long, ByRef blank_stage As Boolean, ByRef allBlank As Boolean) As Long
On Error Resume Next
    
    ''''------------------------------------------------------------------------------
    ''''Definition of the Variables
    ''''------------------------------------------------------------------------------
    ''''       inWave: eFuse DSSC Capture Data Wave
    '''' bitFlag_mode: 0=early_stage, [to decide which bitFlagWave(Stage_early or Stage)]
    ''''               1=normal stage,
    ''''              >1= all bits
    ''''------------------------------------------------------------------------------
    ''''       fbcSum: eFuse Fail Bit Count (Single vs. Double)
    ''''  blank_stage: eFuse blank check for the current stage/job bits only
    ''''     allblank: eFuse blank check for all bits
    ''''------------------------------------------------------------------------------

    Dim m_size As Long
    Dim sgSum As Long
    Dim dbSum As Long
    Dim m_tmpWave As New DSPWave
    
    Dim m_calcsum1 As Long
    Dim m_tmpWave1 As New DSPWave
    Dim m_stageBitFlagWave As New DSPWave
    Dim m_cmpsgWave As New DSPWave
    Dim m_readCateWave As New DSPWave
    
    ''''m_singleWave: eFuse SingleBit Data
    ''''m_doubleWave: eFuse DoubleBit Data => will be used to do the decode section
    Dim m_singleWave As New DSPWave
    Dim m_doubleWave As New DSPWave
    
    ''''initialize
    allBlank = True
    blank_stage = True
    
    m_tmpWave = InWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)
    m_singleWave = m_tmpWave.ConvertDataTypeTo(DspLong).Copy

    Call eFuse_Gen_DoubleBitWave(m_singleWave, m_doubleWave, sgSum, dbSum, fbcSum, m_cmpsgWave)
    m_size = m_doubleWave.SampleSize
    ''''201811XX update
    Call eFuse_decode_DSSCReadWave(FuseType, m_doubleWave, m_readCateWave)
    
    ''''check the blank bits
    If (sgSum <> 0) Then allBlank = False

''''''''-----------------------------------
''''''''From Enum eFuseBlockType
''''    eFuse_ECID = 1
''''    eFuse_CFG = 2
''''    eFuse_UID = 3
''''    eFuse_SEN = 4
''''    eFuse_MON = 5
''''    eFuse_UDR = 6
''''    eFuse_UDRE = 7
''''    eFuse_UDRP = 8
''''    eFuse_CMP = 9
''''    eFuse_CMPE = 10
''''    eFuse_CMPP = 11
''''    eFuse_Block_Unknown = 999
''''''''-----------------------------------

    If (FuseType = 1) Then
        gDW_ECID_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_ECID_Read_SingleBitWave = m_singleWave.Copy
        gDW_ECID_Read_DoubleBitWave = m_doubleWave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_ECID_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_ECID_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 2) Then
        gDW_CFG_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_CFG_Read_SingleBitWave = m_singleWave.Copy
        gDW_CFG_Read_DoubleBitWave = m_doubleWave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_CFG_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_CFG_Stage_BitFlag.Copy
        ElseIf (bitFlag_mode = 3) Then
            m_stageBitFlagWave = gDW_CFG_Stage_Real_BitFlag.Copy
        End If
    ElseIf (FuseType = 3) Then
        gDW_UID_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_UID_Read_SingleBitWave = m_singleWave.Copy
        gDW_UID_Read_DoubleBitWave = m_doubleWave.Copy
        m_stageBitFlagWave = gDW_UID_Stage_BitFlag.Copy
    ElseIf (FuseType = 4) Then
        gDW_SEN_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_SEN_Read_SingleBitWave = m_singleWave.Copy
        gDW_SEN_Read_DoubleBitWave = m_doubleWave.Copy
        m_stageBitFlagWave = gDW_SEN_Stage_BitFlag.Copy
    ElseIf (FuseType = 5) Then
        gDW_MON_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_MON_Read_SingleBitWave = m_singleWave.Copy
        gDW_MON_Read_DoubleBitWave = m_doubleWave.Copy
        m_stageBitFlagWave = gDW_MON_Stage_BitFlag.Copy
    ElseIf (FuseType = 6) Then
    ElseIf (FuseType = 7) Then
    ElseIf (FuseType = 8) Then
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If

    'If (bitFlag_mode > 1) Then
    If (bitFlag_mode > 1 And bitFlag_mode <> 3) Then
        ''''use all bits
        m_stageBitFlagWave.CreateConstant 1, m_size, DspLong
    End If
    m_tmpWave1 = m_doubleWave.bitwiseand(m_stageBitFlagWave)
    m_calcsum1 = m_tmpWave1.CalcSum
    If (m_calcsum1 <> 0) Then blank_stage = False

End Function

Public Function eFuse_Wave1bit_to_SingleDoubleBitWave(ByVal FuseType As Long, ByVal bitFlag_mode As Long, _
                                                      ByVal caseFlag As Boolean, ByVal InWave As DSPWave, _
                                                      ByRef fbcSum As Long, ByRef blank_stage As Boolean, ByRef allBlank As Boolean) As Long
On Error Resume Next
    
    ''''------------------------------------------------------------------------------
    ''''Definition of the Variables
    ''''------------------------------------------------------------------------------
    ''''  inWave: eFuse DSSC Capture Data Wave
    ''''out1Wave: eFuse SingleBit Data
    ''''out2Wave: eFuse DoubleBit Data => will be used to do the decode section
    ''''  fbcSum: eFuse Fail Bit Count (Single vs. Double)
    ''''allblank: eFuse blank check for all bits
    ''''------------------------------------------------------------------------------
    
    Dim m_size As Long
    Dim sgSum As Long
    Dim dbSum As Long
    Dim m_tmpWave As New DSPWave
    Dim m_cmpsgWave As New DSPWave
    Dim m_readCateWave As New DSPWave
    Dim m_stageBitFlagWave As New DSPWave
    Dim m_CalcSum As Long
    Dim out1Wave As New DSPWave
    Dim out2Wave As New DSPWave
    

    If (caseFlag = False) Then
        ''''bit0_bitLast
        Call eFuse_DspWave_Copy(InWave, out1Wave)
    Else
        ''''bitLast_bit0
        Call eFuse_reverseBitWave(InWave, out1Wave)
    End If
    
    Call eFuse_Gen_DoubleBitWave(out1Wave, out2Wave, sgSum, dbSum, fbcSum, m_cmpsgWave)
    m_size = out2Wave.SampleSize
    
    ''''201811XX update
    Call eFuse_decode_DSSCReadWave(FuseType, out2Wave, m_readCateWave)
    
    allBlank = True ''''initial
    If (sgSum <> 0) Then allBlank = False

    If (FuseType = 1) Then
        gDW_ECID_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_ECID_Read_SingleBitWave = out1Wave.Copy
        gDW_ECID_Read_DoubleBitWave = out2Wave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_ECID_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_ECID_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 2) Then
        gDW_CFG_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_CFG_Read_SingleBitWave = out1Wave.Copy
        gDW_CFG_Read_DoubleBitWave = out2Wave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_CFG_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_CFG_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 3) Then
    ElseIf (FuseType = 4) Then
    ElseIf (FuseType = 5) Then
        gDW_MON_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_MON_Read_SingleBitWave = out1Wave.Copy
        gDW_MON_Read_DoubleBitWave = out2Wave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_MON_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_MON_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 6) Then
        gDW_UDR_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_UDR_Read_SingleBitWave = out1Wave.Copy
        gDW_UDR_Read_DoubleBitWave = out2Wave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_UDR_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_UDR_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 7) Then
        gDW_UDRE_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_UDRE_Read_SingleBitWave = out1Wave.Copy
        gDW_UDRE_Read_DoubleBitWave = out2Wave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_UDRE_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_UDRE_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 8) Then
        gDW_UDRP_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_UDRP_Read_SingleBitWave = out1Wave.Copy
        gDW_UDRP_Read_DoubleBitWave = out2Wave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_UDRP_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_UDRP_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 9) Then
        gDW_CMP_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_CMP_Read_SingleBitWave = out1Wave.Copy
        gDW_CMP_Read_DoubleBitWave = out2Wave.Copy
'        If (bitFlag_mode = 0) Then
'            m_stageBitFlagWave = gDW_UDRP_Stage_Early_BitFlag.Copy
'        ElseIf (bitFlag_mode = 1) Then
        If (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_CMP_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 10) Then
        gDW_CMPE_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_CMPE_Read_SingleBitWave = out1Wave.Copy
        gDW_CMPE_Read_DoubleBitWave = out2Wave.Copy
'        If (bitFlag_mode = 0) Then
'            m_stageBitFlagWave = gDW_UDRP_Stage_Early_BitFlag.Copy
'        ElseIf (bitFlag_mode = 1) Then
        If (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_CMPE_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 11) Then
        gDW_CMPP_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_CMPP_Read_SingleBitWave = out1Wave.Copy
        gDW_CMPP_Read_DoubleBitWave = out2Wave.Copy
'        If (bitFlag_mode = 0) Then
'            m_stageBitFlagWave = gDW_UDRP_Stage_Early_BitFlag.Copy
'        ElseIf (bitFlag_mode = 1) Then
        If (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_CMPP_Stage_BitFlag.Copy
        End If
    Else
    End If
    
    If (bitFlag_mode > 1) Then
        ''''use all bits
        m_stageBitFlagWave.CreateConstant 1, m_size, DspLong
    End If
    m_tmpWave = out2Wave.bitwiseand(m_stageBitFlagWave)
    m_CalcSum = m_tmpWave.CalcSum
    If (m_CalcSum <> 0) Then blank_stage = False

End Function

Public Function eFuse_Gen_PgmBitSrcWave(ByVal FuseType As Long, ByVal bitFlag_mode As Long, _
                                        ByRef outSrcWave As DSPWave, ByRef result As Long) As Long
On Error Resume Next

    '''' bitFlag_mode: 0=early_stage, 1=normal stage to decide which bitFlagWave(Stage_early or Stage)
    
    Dim i As Long, j As Long, k As Long
    Dim k1 As Long, k2 As Long
    Dim m_tmpValue As Long
    Dim m_sgbits As Long
    
    Dim expandWidth As Long
    Dim BitsPerRow As Long
    Dim ReadCycles As Long
    Dim BitsPerCycle As Long
    Dim BitsPerBlock As Long
    
    Dim m_tmpWave1 As New DSPWave
    Dim m_tmpWave2 As New DSPWave
    Dim m_pgmrawBitWave As New DSPWave
    Dim m_defaultBitWave As New DSPWave
    Dim m_effbitFlagWave As New DSPWave
    Dim m_singleWave As New DSPWave
    Dim m_doubleWave As New DSPWave
    
    Dim m_size As Long
    Dim m_tmpArr() As Long
    Dim m_dbbitArr() As Long
    Dim m_sgBitArr() As Long
    Dim m_outSrcBitArr() As Long
    
    BitsPerRow = gDL_BitsPerRow
    ReadCycles = gDL_ReadCycles
    BitsPerCycle = gDL_BitsPerCycle
    BitsPerBlock = gDL_BitsPerBlock
    expandWidth = gDL_DigSrcRepeatN

''''''''-----------------------------------
''''''''From Enum eFuseBlockType
''''    eFuse_ECID = 1
''''    eFuse_CFG = 2
''''    eFuse_UID = 3
''''    eFuse_SEN = 4
''''    eFuse_MON = 5
''''    eFuse_UDR = 6
''''    eFuse_UDRE = 7
''''    eFuse_UDRP = 8
''''    eFuse_CMP = 9
''''    eFuse_CMPE = 10
''''    eFuse_CMPP = 11
''''    eFuse_Block_Unknown = 999
''''''''-----------------------------------
    
    If (FuseType = 1) Then
        m_defaultBitWave = gDW_ECID_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_ECID_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_ECID_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 2) Then
        m_defaultBitWave = gDW_CFG_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_CFG_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_CFG_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 3) Then
        m_defaultBitWave = gDW_UID_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_UID_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_UID_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 4) Then
        m_defaultBitWave = gDW_SEN_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_SEN_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_SEN_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 5) Then
        m_defaultBitWave = gDW_MON_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_MON_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_MON_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 6) Then
        m_defaultBitWave = gDW_UDR_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_UDR_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_UDR_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 7) Then
        m_defaultBitWave = gDW_UDRE_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_UDRE_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_UDRE_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 8) Then
        m_defaultBitWave = gDW_UDRP_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_UDRP_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_UDRP_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If

    If (bitFlag_mode > 1) Then
        ''''use all bits
        m_effbitFlagWave.CreateConstant 1, BitsPerBlock, DspLong
    End If

    m_pgmrawBitWave = gDW_Pgm_RawBitWave.Copy

    ''''combine m_pgmrawBitWave with "OR" the m_defaultBitWave
    m_tmpWave1 = m_pgmrawBitWave.BitwiseOr(m_defaultBitWave).Copy
    
    ''''gen effective Wave with "AND" the effbitFlagWave
    m_doubleWave = m_tmpWave1.bitwiseand(m_effbitFlagWave).Copy

    m_sgbits = BitsPerCycle * ReadCycles
    m_singleWave.CreateConstant 0, m_sgbits, DspLong
    
    If (gDL_eFuse_Orientation = 0) Then ''''Up2Down
        m_singleWave = m_doubleWave.repeat(2).Copy

    ElseIf (gDL_eFuse_Orientation = 1) Then ''''SingleUp, eFuse_1_Bit
        ''''doubleBitWave is equal to singleBitWave
        m_singleWave = m_doubleWave.Copy

    ElseIf (gDL_eFuse_Orientation = 2) Then ''''Right2Left, eFuse_2_Bit
        ''''-------------------------------------------------------------------------------
        ''''New Method
        ''''-------------------------------------------------------------------------------
        m_dbbitArr = m_doubleWave.Data
        m_sgBitArr = m_singleWave.Data

        k = 0 ''''must be here
        For i = 0 To ReadCycles - 1     'EX: EcidReadCycle - 1   ''0...15(ECID)
            For j = 0 To BitsPerRow - 1 'EX: EcidBitsPerRow - 1  ''0...15(ECID)
                ''''k1: Right block
                ''''k2:  Left block
                k1 = (i * BitsPerCycle) + j ''<Important> Must use BitsPerCycle here
                k2 = (i * BitsPerCycle) + BitsPerRow + j
                ''m_tmpValue = m_doubleWave.Element(k)
                ''m_singleWave.Element(k1) = m_tmpValue
                ''m_singleWave.Element(k2) = m_tmpValue
                
                ''''save TT here by using dataArr
                m_tmpValue = m_dbbitArr(k)
                m_sgBitArr(k1) = m_tmpValue
                m_sgBitArr(k2) = m_tmpValue
                k = k + 1
            Next j
        Next i
        m_singleWave.Data = m_sgBitArr ''''save TT
        ''''-------------------------------------------------------------------------------

    ''''Below is reserved for the future, there is NO any definition at present.
    ElseIf (gDL_eFuse_Orientation = 3) Then
    ElseIf (gDL_eFuse_Orientation = 4) Then
    ElseIf (gDL_eFuse_Orientation = 5) Then
    End If

    ''''Expand the outWave to Source Wave
    m_size = m_sgbits * expandWidth
    'outSrcWave.CreateConstant 0, m_size, DspLong
    outSrcWave.CreateConstant 0, m_sgbits, DspLong
    
    ReDim m_outSrcBitArr(m_size - 1)
    
    If (gDL_eFuse_Orientation <> 2) Then m_sgBitArr = m_singleWave.Data
    
'    k = 0
'    For i = 0 To m_sgbits - 1
'        For j = 0 To expandWidth - 1
'            k = i * expandWidth + j
'            ''outSrcWave.Element(k) = m_singleWave.Element(i) ''''waste TT
'            m_outSrcBitArr(k) = m_sgBitArr(i) ''''to save TT
'        Next j
'    Next i
'    outSrcWave.Data = m_outSrcBitArr ''''to save TT
    outSrcWave.Data = m_sgBitArr
    
    ''''---------------------------------------------------------------
    ''''Update to global DSP variables
    ''''---------------------------------------------------------------
    If (FuseType = 1) Then
        gDW_ECID_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_ECID_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 2) Then
        gDW_CFG_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_CFG_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 3) Then
        gDW_UID_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_UID_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 4) Then
        gDW_SEN_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_SEN_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 5) Then
        gDW_MON_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_MON_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 6) Then
        gDW_UDR_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_UDR_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 7) Then
        gDW_UDRE_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_UDRE_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 8) Then
        gDW_UDRP_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_UDRP_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If
    ''''---------------------------------------------------------------

    result = 1

End Function

Public Function eFuse_Gen_SingleDoubleWave(ByVal InWave As DSPWave, ByVal defaultBitWave As DSPWave, ByVal effbitFlagWave As DSPWave, _
                                           ByRef doubleWave As DSPWave, ByRef singleWave As DSPWave) As Long
On Error Resume Next

    Dim i As Long, j As Long, k As Long
    Dim k1 As Long, k2 As Long
    Dim m_tmpValue As Long
    Dim m_pgmbits As Long
    
    Dim BitsPerRow As Long
    Dim ReadCycles As Long
    Dim BitsPerCycle As Long
    Dim BitsPerBlock As Long

    Dim m_dbbitArr() As Long
    Dim m_sgBitArr() As Long
    Dim m_tmpWave1 As New DSPWave
    
    BitsPerRow = gDL_BitsPerRow
    ReadCycles = gDL_ReadCycles
    BitsPerCycle = gDL_BitsPerCycle
    BitsPerBlock = gDL_BitsPerBlock

    ''''combine inWave with "OR" the defaultBitWave
    m_tmpWave1.CreateConstant 0, InWave.SampleSize, DspLong
    m_tmpWave1 = InWave.BitwiseOr(defaultBitWave).Copy
    
    ''''gen effective Wave with "AND" the effbitFlagWave
    doubleWave.CreateConstant 0, InWave.SampleSize, DspLong
    doubleWave = m_tmpWave1.bitwiseand(effbitFlagWave).Copy

    m_pgmbits = BitsPerCycle * ReadCycles
    singleWave.CreateConstant 0, m_pgmbits, DspLong
    
    If (gDL_eFuse_Orientation = 0) Then ''''Up2Down
        singleWave = doubleWave.repeat(2).Copy

    ElseIf (gDL_eFuse_Orientation = 1) Then ''''SingleUp, eFuse_1_Bit
        ''''doubleBitWave is equal to singleBitWave
        singleWave = doubleWave.Copy

    ElseIf (gDL_eFuse_Orientation = 2) Then ''''Right2Left, eFuse_2_Bit
        ''''-------------------------------------------------------------------------------
        ''''New Method
        ''''-------------------------------------------------------------------------------
        m_dbbitArr = doubleWave.Data
        m_sgBitArr = singleWave.Data

        k = 0 ''''must be here
        For i = 0 To ReadCycles - 1     'EX: EcidReadCycle - 1   ''0...15(ECID)
            For j = 0 To BitsPerRow - 1 'EX: EcidBitsPerRow - 1  ''0...15(ECID)
                ''''k1: Right block
                ''''k2:  Left block
                k1 = (i * BitsPerCycle) + j ''<Important> Must use BitsPerCycle here
                k2 = (i * BitsPerCycle) + BitsPerRow + j
                ''m_tmpValue = doubleWave.Element(k)
                ''singleWave.Element(k1) = m_tmpValue
                ''singleWave.Element(k2) = m_tmpValue
                
                ''''save TT here by using dataArr
                m_tmpValue = m_dbbitArr(k)
                m_sgBitArr(k1) = m_tmpValue
                m_sgBitArr(k2) = m_tmpValue
                k = k + 1
            Next j
        Next i
        singleWave.Data = m_sgBitArr ''''save TT
        ''''-------------------------------------------------------------------------------

    ''''Below is reserved for the future, there is NO any definition at present.
    ElseIf (gDL_eFuse_Orientation = 3) Then
    ElseIf (gDL_eFuse_Orientation = 4) Then
    ElseIf (gDL_eFuse_Orientation = 5) Then
    End If

End Function

Public Function eFuse_singleWave_to_32Bits_CapWave(ByVal singleWave As DSPWave, ByRef outcapWave As DSPWave) As Long
On Error Resume Next

    Dim m_tmpWave As New DSPWave
    
    ''''Convert singleWave to the 32-bit CapWave
    m_tmpWave = singleWave.ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb).ConvertNumFormatTo(TwosComplement).ConvertDataTypeTo(DspDouble)
    outcapWave = m_tmpWave.Copy

End Function

''''Here inWave as PgmBitWave or DoubleBitWave
Public Function eFuse_Sim_Gen_32Bits_CapWave(ByVal FuseType As Long, ByVal simBlank As Long, _
                                             ByRef outcapWave As DSPWave, ByVal Reverse As Boolean) As Long
On Error Resume Next

    ''''------------------------------------------------------------------------------
    ''''Definition of the Variables
    ''''------------------------------------------------------------------------------
    ''''  simBlank: simulate blank condition to decide which stage bit flag to be used
    ''''       = 0: means that all bits blank=True as early stage bits
    ''''       = 1: means that simulate those bits (stage <  job)
    ''''       = 2: means that simulate those bits (stage <= job)
    ''''------------------------------------------------------------------------------
    
    Dim i As Long, j As Long, k As Long
    Dim k1 As Long, k2 As Long
    Dim m_tmpValue As Long
    Dim m_pgmbits As Long
    
    Dim BitsPerRow As Long
    Dim ReadCycles As Long
    Dim BitsPerCycle As Long
    Dim BitsPerBlock As Long

    Dim m_dbbitArr() As Long
    Dim m_sgBitArr() As Long
    Dim m_tmpWave As New DSPWave
    Dim m_tmpWave1 As New DSPWave
    
    Dim m_inWave As New DSPWave
    Dim m_defaultBitWave As New DSPWave
    Dim m_effbitFlagWave As New DSPWave
    Dim m_singleWave As New DSPWave
    Dim m_doubleWave As New DSPWave
    
    BitsPerRow = gDL_BitsPerRow
    ReadCycles = gDL_ReadCycles
    BitsPerCycle = gDL_BitsPerCycle
    BitsPerBlock = gDL_BitsPerBlock

    m_pgmbits = BitsPerCycle * ReadCycles
    
''''''''-----------------------------------
''''''''From Enum eFuseBlockType
''''    eFuse_ECID = 1
''''    eFuse_CFG = 2
''''    eFuse_UID = 3
''''    eFuse_SEN = 4
''''    eFuse_MON = 5
''''    eFuse_UDR = 6
''''    eFuse_UDRE = 7
''''    eFuse_UDRP = 8
''''    eFuse_CMP = 9
''''    eFuse_CMPE = 10
''''    eFuse_CMPP = 11
''''    eFuse_Block_Unknown = 999
''''''''-----------------------------------

    If (FuseType = 1) Then
        m_defaultBitWave = gDW_ECID_allDefaultBitWave.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_ECID_StageLEQJob_BitFlag.Subtract(gDW_ECID_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_ECID_StageLEQJob_BitFlag.Copy
        End If
    ElseIf (FuseType = 2) Then
        m_defaultBitWave = gDW_CFG_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Subtract(gDW_CFG_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        End If
    ElseIf (FuseType = 3) Then
        m_defaultBitWave = gDW_UID_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_UID_StageLEQJob_BitFlag.Subtract(gDW_UID_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_UID_StageLEQJob_BitFlag.Copy
        End If
    ElseIf (FuseType = 4) Then
        m_defaultBitWave = gDW_SEN_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_SEN_StageLEQJob_BitFlag.Subtract(gDW_SEN_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_SEN_StageLEQJob_BitFlag.Copy
        End If
    ElseIf (FuseType = 5) Then
        m_defaultBitWave = gDW_MON_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_MON_StageLEQJob_BitFlag.Subtract(gDW_MON_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_MON_StageLEQJob_BitFlag.Copy
        End If
    ElseIf (FuseType = 6) Then
        m_defaultBitWave = gDW_UDR_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_UDR_StageLEQJob_BitFlag.Subtract(gDW_UDR_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_UDR_StageLEQJob_BitFlag.Copy
        End If
        'm_pgmbits = gL_USI_DigSrcBits_Num
    ElseIf (FuseType = 7) Then
        m_defaultBitWave = gDW_UDRE_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_UDRE_StageLEQJob_BitFlag.Subtract(gDW_UDRE_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_UDRE_StageLEQJob_BitFlag.Copy
        End If
    ElseIf (FuseType = 8) Then
        m_defaultBitWave = gDW_UDRP_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_UDRP_StageLEQJob_BitFlag.Subtract(gDW_UDRP_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_UDRP_StageLEQJob_BitFlag.Copy
        End If
    ElseIf (FuseType = 9) Then
        m_defaultBitWave = gDW_CMP_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_CMP_StageLEQJob_BitFlag.Subtract(gDW_CMP_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_CMP_StageLEQJob_BitFlag.Copy
        End If
    ElseIf (FuseType = 10) Then
        m_defaultBitWave = gDW_CMPE_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_CMPE_StageLEQJob_BitFlag.Subtract(gDW_CMPE_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_CMPE_StageLEQJob_BitFlag.Copy
        End If
    ElseIf (FuseType = 11) Then
        m_defaultBitWave = gDW_CMPP_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy
        If (simBlank = 0) Then
            m_effbitFlagWave.CreateConstant 0, m_defaultBitWave.SampleSize, DspLong
        ElseIf (simBlank = 1) Then
            m_effbitFlagWave = gDW_CMPP_StageLEQJob_BitFlag.Subtract(gDW_CMPP_Stage_BitFlag).Copy
        ElseIf (simBlank = 2) Then
            m_effbitFlagWave = gDW_CMPP_StageLEQJob_BitFlag.Copy
        End If
    Else
    End If

    m_inWave = gDW_Pgm_RawBitWave.Copy
    
    ''''combine inWave with "OR" the defaultBitWave
    ''m_tmpWave1.CreateConstant 0, m_inWave.SampleSize, DspLong
    m_tmpWave1 = m_inWave.BitwiseOr(m_defaultBitWave).Copy
    
    ''''gen effective Wave with "AND" the effbitFlagWave
    m_doubleWave.CreateConstant 0, m_inWave.SampleSize, DspLong
    m_doubleWave = m_tmpWave1.bitwiseand(m_effbitFlagWave).Copy

    
    m_singleWave.CreateConstant 0, m_pgmbits, DspLong
    
    If (gDL_eFuse_Orientation = 0) Then ''''Up2Down
        m_singleWave = m_doubleWave.repeat(2).Copy

    ElseIf (gDL_eFuse_Orientation = 1) Then ''''SingleUp, eFuse_1_Bit
        ''''doubleBitWave is equal to singleBitWave
        m_singleWave = m_doubleWave.Copy

    ElseIf (gDL_eFuse_Orientation = 2) Then ''''Right2Left, eFuse_2_Bit
        ''''-------------------------------------------------------------------------------
        ''''New Method
        ''''-------------------------------------------------------------------------------
        m_dbbitArr = m_doubleWave.Data
        m_sgBitArr = m_singleWave.Data
        
        k = 0 ''''must be here
        For i = 0 To ReadCycles - 1     'EX: EcidReadCycle - 1   ''0...15(ECID)
            For j = 0 To BitsPerRow - 1 'EX: EcidBitsPerRow - 1  ''0...15(ECID)
                ''''k1: Right block
                ''''k2:  Left block
                k1 = (i * BitsPerCycle) + j ''<Important> Must use BitsPerCycle here
                k2 = (i * BitsPerCycle) + BitsPerRow + j

                ''''save TT here by using dataArr
                m_tmpValue = m_dbbitArr(k)
                m_sgBitArr(k1) = m_tmpValue
                m_sgBitArr(k2) = m_tmpValue
                k = k + 1
            Next j
        Next i
        m_singleWave.Data = m_sgBitArr ''''save TT
        ''''-------------------------------------------------------------------------------

    ''''Below is reserved for the future, there is NO any definition at present.
    ElseIf (gDL_eFuse_Orientation = 3) Then
    ElseIf (gDL_eFuse_Orientation = 4) Then
    ElseIf (gDL_eFuse_Orientation = 5) Then
    End If

    ''''Convert singleWave to the 32-bit CapWave
    m_tmpWave.CreateConstant 0, ReadCycles, DspDouble
    'If (FuseType >= 6) Then
    If (gDB_SerialType = True) Then ''20190801
        m_tmpWave = m_singleWave.Copy
        If (Reverse = True) Then Call eFuse_reverseBitWave(m_tmpWave, outcapWave)
    Else
        m_tmpWave = m_singleWave.ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb).ConvertNumFormatTo(TwosComplement).ConvertDataTypeTo(DspDouble)
        outcapWave = m_tmpWave.Copy
    End If
    

End Function

Public Function eFuse_updatePgmWave_byCategory(ByVal indexWave As DSPWave, ByVal BitArrWave As DSPWave) As Long
On Error Resume Next

    Dim m_tmpWave As New DSPWave

    m_tmpWave = gDW_Pgm_RawBitWave.Copy
    Call m_tmpWave.ReplaceElements(indexWave, BitArrWave)
    gDW_Pgm_RawBitWave = m_tmpWave.Copy

End Function

Public Function eFuse_compare_Read_PgmBitWave(ByVal FuseType As Long, ByVal bitFlag_mode As Long, ByVal InWave As DSPWave, _
                                              ByRef fbcSum As Long, ByRef cmpResult As Long, ByVal ReverseFlag As Boolean) As Long
On Error Resume Next

    ''''------------------------------------------------------------------------------
    ''''Definition of the Variables
    ''''------------------------------------------------------------------------------
    ''''       inWave: eFuse DSSC Capture Data Wave
    ''''------------------------------------------------------------------------------
    ''''       inWave: eFuse DSSC Capture Data Wave
    '''' bitFlag_mode: 0=early_stage, [to decide which bitFlagWave(Stage_early or Stage)]
    ''''               1=normal stage,
    ''''              >1= all bits
    ''''------------------------------------------------------------------------------
    ''''       fbcSum: eFuse Read Fail Bit Count (Single vs. Double)
    ''''    cmpResult: 0 means Equal (Read=Pgm), >0 means Not Equal (Read<>Pgm)
    ''''------------------------------------------------------------------------------

    Dim m_size As Long
    Dim sgSum As Long
    Dim dbSum As Long
    Dim m_pgmBitWave As New DSPWave
    Dim m_cmpResWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_tmpWave1 As New DSPWave
    Dim m_stageBitFlagWave As New DSPWave
    Dim m_cmpsgWave As New DSPWave
    Dim m_readCate As New DSPWave
    Dim m_sgBitWave As New DSPWave
    Dim m_dbBitWave As New DSPWave
    
''''''''-----------------------------------
''''''''From Enum eFuseBlockType
''''    eFuse_ECID = 1
''''    eFuse_CFG = 2
''''    eFuse_UID = 3
''''    eFuse_SEN = 4
''''    eFuse_MON = 5
''''    eFuse_UDR = 6
''''    eFuse_UDRE = 7
''''    eFuse_UDRP = 8
''''    eFuse_CMP = 9
''''    eFuse_CMPE = 10
''''    eFuse_CMPP = 11
''''    eFuse_Block_Unknown = 999
''''''''-----------------------------------

    If (FuseType = 1) Then
        m_pgmBitWave = gDW_ECID_Pgm_DoubleBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_ECID_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_ECID_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 2) Then
        m_pgmBitWave = gDW_CFG_Pgm_DoubleBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_CFG_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_CFG_Stage_BitFlag.Copy
        ElseIf (bitFlag_mode = 3) Then
            m_stageBitFlagWave = gDW_CFG_Stage_Real_BitFlag.Copy
        End If
    ElseIf (FuseType = 3) Then
        m_pgmBitWave = gDW_UID_Pgm_DoubleBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_UID_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_UID_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 4) Then
        m_pgmBitWave = gDW_SEN_Pgm_DoubleBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_SEN_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_SEN_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 5) Then
        m_pgmBitWave = gDW_MON_Pgm_DoubleBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_MON_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_MON_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 6) Then
        m_pgmBitWave = gDW_UDR_Pgm_DoubleBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_UDR_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_UDR_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 7) Then
        m_pgmBitWave = gDW_UDRE_Pgm_DoubleBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_UDRE_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_UDRE_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 8) Then
        m_pgmBitWave = gDW_UDRP_Pgm_DoubleBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_stageBitFlagWave = gDW_UDRP_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_stageBitFlagWave = gDW_UDRP_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If

    If (bitFlag_mode > 1 And bitFlag_mode <> 3) Then
        m_size = m_pgmBitWave.SampleSize
        m_stageBitFlagWave.CreateConstant 1, m_size, DspLong
    End If
        
    ''''initialize
'    m_tmpWave = inWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)
'    m_sgBitWave = m_tmpWave.ConvertDataTypeTo(DspLong).Copy
    If (gDB_SerialType = False) Then ''21090801
        m_tmpWave = InWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)
        m_sgBitWave = m_tmpWave.ConvertDataTypeTo(DspLong).Copy
    Else
        Dim out1Wave As New DSPWave
        If (ReverseFlag = False) Then
         ''''bit0_bitLast
            Call eFuse_DspWave_Copy(InWave, out1Wave)
        Else
            ''''bitLast_bit0
            Call eFuse_reverseBitWave(InWave, out1Wave)
        End If
        m_sgBitWave = out1Wave.ConvertDataTypeTo(DspLong).Copy
    End If

    Call eFuse_Gen_DoubleBitWave(m_sgBitWave, m_dbBitWave, sgSum, dbSum, fbcSum, m_cmpsgWave)
    
    ''''201811XX update
    Call eFuse_decode_DSSCReadWave(FuseType, m_dbBitWave, m_readCate)
    
    ''''Here m_tmpWave1 means the related bits with bitFlagWave
    m_tmpWave1 = m_dbBitWave.ConvertDataTypeTo(DspLong).bitwiseand(m_stageBitFlagWave.ConvertDataTypeTo(DspLong))
    'm_tmpWave1 = m_dbBitWave.BitwiseAnd(m_stageBitFlagWave)

    m_cmpResWave = m_tmpWave1.LogicalCompare(NotEqualTo, m_pgmBitWave.ConvertDataTypeTo(DspLong))
    
    ''''Here LogicalCompare(NotEqualTo) return 1, otherwise return 0 (equal) for each element
    ''''cmpResult means how many bits are NOt equal
    ''''if (cmpResult=0), means Read==PgmBits
    cmpResult = m_cmpResWave.CalcSum

    If (FuseType = 1) Then
        gDW_ECID_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_ECID_Read_SingleBitWave = m_sgBitWave.Copy
        gDW_ECID_Read_DoubleBitWave = m_dbBitWave.Copy
    ElseIf (FuseType = 2) Then
        gDW_CFG_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_CFG_Read_SingleBitWave = m_sgBitWave.Copy
        gDW_CFG_Read_DoubleBitWave = m_dbBitWave.Copy
    ElseIf (FuseType = 3) Then
        gDW_UID_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_UID_Read_SingleBitWave = m_sgBitWave.Copy
        gDW_UID_Read_DoubleBitWave = m_dbBitWave.Copy
    ElseIf (FuseType = 4) Then
        gDW_SEN_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_SEN_Read_SingleBitWave = m_sgBitWave.Copy
        gDW_SEN_Read_DoubleBitWave = m_dbBitWave.Copy
    ElseIf (FuseType = 5) Then
        gDW_MON_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_MON_Read_SingleBitWave = m_sgBitWave.Copy
        gDW_MON_Read_DoubleBitWave = m_dbBitWave.Copy
    ElseIf (FuseType = 6) Then
        gDW_UDR_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_UDR_Read_SingleBitWave = m_sgBitWave.Copy
        gDW_UDR_Read_DoubleBitWave = m_dbBitWave.Copy
    ElseIf (FuseType = 7) Then
        gDW_UDRE_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_UDRE_Read_SingleBitWave = m_sgBitWave.Copy
        gDW_UDRE_Read_DoubleBitWave = m_dbBitWave.Copy
    ElseIf (FuseType = 8) Then
        gDW_UDRP_Read_cmpsgWavePerCyc = m_cmpsgWave.Copy
        gDW_UDRP_Read_SingleBitWave = m_sgBitWave.Copy
        gDW_UDRP_Read_DoubleBitWave = m_dbBitWave.Copy
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If

End Function

''''201811XX update, it is used to do the simulation.
Public Function eFuse_SingleBitWave2CapWave32Bits(ByVal FuseType As Long, ByRef outcapWave As DSPWave) As Long
On Error Resume Next

    Dim m_tmpWave1 As New DSPWave
    Dim m_tmpWave2 As New DSPWave
    Dim m_pgmSgWave As New DSPWave
    Dim m_readSgWave As New DSPWave

    If (FuseType = 1) Then
        m_pgmSgWave = gDW_ECID_Pgm_SingleBitWave.Copy
        m_readSgWave = gDW_ECID_Read_SingleBitWave.Copy
    ElseIf (FuseType = 2) Then
        m_pgmSgWave = gDW_CFG_Pgm_SingleBitWave.Copy
        m_readSgWave = gDW_CFG_Read_SingleBitWave.Copy
    ElseIf (FuseType = 3) Then
        m_pgmSgWave = gDW_UID_Pgm_SingleBitWave.Copy
        m_readSgWave = gDW_UID_Read_SingleBitWave.Copy
    ElseIf (FuseType = 4) Then
        m_pgmSgWave = gDW_SEN_Pgm_SingleBitWave.Copy
        m_readSgWave = gDW_SEN_Read_SingleBitWave.Copy
    ElseIf (FuseType = 5) Then
        m_pgmSgWave = gDW_MON_Pgm_SingleBitWave.Copy
        m_readSgWave = gDW_MON_Read_SingleBitWave.Copy
    ElseIf (FuseType = 6) Then
        m_pgmSgWave = gDW_UDR_Pgm_SingleBitWave.Copy
        m_readSgWave = gDW_UDR_Read_SingleBitWave.Copy
    ElseIf (FuseType = 7) Then
        m_pgmSgWave = gDW_UDRE_Pgm_SingleBitWave.Copy
        m_readSgWave = gDW_UDRE_Read_SingleBitWave.Copy
    ElseIf (FuseType = 8) Then
        m_pgmSgWave = gDW_UDRP_Pgm_SingleBitWave.Copy
        m_readSgWave = gDW_UDRP_Read_SingleBitWave.Copy
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If

    ''''combine inWave with "OR" the defaultBitWave
    ''outsingleWave.CreateConstant 0, m_size, DspLong
    m_pgmSgWave = m_pgmSgWave.ConvertNumFormatTo(TwosComplement).ConvertDataTypeTo(DspLong).Copy
    m_readSgWave = m_readSgWave.ConvertNumFormatTo(TwosComplement).ConvertDataTypeTo(DspLong).Copy
    m_tmpWave1 = m_pgmSgWave.BitwiseOr(m_readSgWave).Copy

    ''''Convert singleWave to the 32-bit CapWave
    ''outcapWave.CreateConstant 0, ReadCycles, DspDouble
    ''m_tmpWave2.CreateConstant 0, ReadCycles, DspDouble
    If (gDL_eFuse_Orientation = 1) Then
        outcapWave = m_tmpWave1.Copy
    ElseIf (gDL_eFuse_Orientation = 2) Then
'        m_tmpWave2 = m_tmpWave1.ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb)
'        outcapWave = m_tmpWave2.ConvertNumFormatTo(TwosComplement).ConvertDataTypeTo(DspDouble).Copy
        If (gDB_SerialType = False) Then
            m_tmpWave2 = m_tmpWave1.ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb)
            outcapWave = m_tmpWave2.ConvertNumFormatTo(TwosComplement).ConvertDataTypeTo(DspDouble).Copy
        Else
            outcapWave = m_tmpWave1.Copy
        End If
    End If
    'outcapWave = m_tmpWave2.ConvertNumFormatTo(TwosComplement).ConvertDataTypeTo(DspDouble).Copy
    
    If (FuseType = 1) Then
        gDW_ECID_Read_SingleBitWave = m_tmpWave1.Copy
    ElseIf (FuseType = 2) Then
        gDW_CFG_Read_SingleBitWave = m_tmpWave1.Copy
    ElseIf (FuseType = 3) Then
        gDW_UID_Read_SingleBitWave = m_tmpWave1.Copy
    ElseIf (FuseType = 4) Then
        gDW_SEN_Read_SingleBitWave = m_tmpWave1.Copy
    ElseIf (FuseType = 5) Then
        gDW_MON_Read_SingleBitWave = m_tmpWave1.Copy
    ElseIf (FuseType = 6) Then
        gDW_UDR_Read_SingleBitWave = m_tmpWave1.Copy
    ElseIf (FuseType = 7) Then
        gDW_UDRE_Read_SingleBitWave = m_tmpWave1.Copy
    ElseIf (FuseType = 8) Then
        gDW_UDRP_Read_SingleBitWave = m_tmpWave1.Copy
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If

End Function

''''unused
Public Function eFuse_Get_ValueFromWaveXX(ByVal idx As Long, ByVal readDecWave As DSPWave, _
                                        ByVal StartBit As Long, ByVal stopbit As Long, ByVal doubleWave As DSPWave, _
                                        ByRef outDec As Double, ByRef bitSum As Long, ByRef outBitWave As DSPWave) As Long
On Error Resume Next
                                        
    Dim m_size As Long
    Dim m_tmpArr() As Double
    Dim m_tmpWave As New DSPWave

    m_size = Abs(stopbit - StartBit) + 1
    m_tmpWave = doubleWave.Select(StartBit, 1, m_size).Copy
    Call eFuse_DspWave_Copy(m_tmpWave, outBitWave)
    bitSum = outBitWave.CalcSum
    
    m_tmpArr = readDecWave.Data
    outDec = m_tmpArr(idx)

End Function

Public Function eFuse_Get_ValueFromWave(ByVal FuseType As Long, ByVal idx As Long, ByVal resolution As Double, ByVal calc_mode As Long, _
                                        ByRef outDec As Double, ByRef outVal As Double, _
                                        ByRef outbitSum As Long, ByRef outBitWave As DSPWave) As Long
On Error Resume Next

    ''''------------------------------------------------------------------------------
    ''''Definition of the Variables
    ''''------------------------------------------------------------------------------
    '''' calc_mode: 0=decimal/default or ids,
    ''''            1=vddbin with safe voltage,
    ''''            2=vddbin with bincut,
    ''''            3=base with safe voltage
    ''''------------------------------------------------------------------------------

    Dim m_reverseFlag As Long
    Dim m_msbbitArr() As Long
    Dim m_lsbbitArr() As Long
    Dim m_bitwidthArr() As Long
    Dim m_startbit As Long
    Dim m_bitwidth As Long
    Dim m_tmpArr() As Double
    
    Dim m_tmpWave As New DSPWave
    Dim m_doubleWave As New DSPWave
    Dim m_readCateWave As New DSPWave
    
    If (FuseType = 1) Then
        m_msbbitArr = gDW_ECID_MSBBit_Cate.Data
        m_lsbbitArr = gDW_ECID_LSBBit_Cate.Data
        m_bitwidthArr = gDW_ECID_BitWidth_Cate.Data
        m_readCateWave = gDW_ECID_Read_Decimal_Cate.Copy
        m_doubleWave = gDW_ECID_Read_DoubleBitWave.Copy
    ElseIf (FuseType = 2) Then
        m_msbbitArr = gDW_CFG_MSBBit_Cate.Data
        m_lsbbitArr = gDW_CFG_LSBBit_Cate.Data
        m_bitwidthArr = gDW_CFG_BitWidth_Cate.Data
        m_readCateWave = gDW_CFG_Read_Decimal_Cate.Copy
        m_doubleWave = gDW_CFG_Read_DoubleBitWave.Copy
    ElseIf (FuseType = 3) Then
    ElseIf (FuseType = 4) Then
    ElseIf (FuseType = 5) Then
    ElseIf (FuseType = 6) Then
    ElseIf (FuseType = 7) Then
    ElseIf (FuseType = 8) Then
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If

    ''''get the correct startbit and bitwidth per idx
    m_bitwidth = m_bitwidthArr(idx)
    If (m_msbbitArr(idx) < m_lsbbitArr(idx)) Then
        ''''ECID case
        m_startbit = m_msbbitArr(idx)
        m_reverseFlag = 1
    Else
        m_startbit = m_lsbbitArr(idx)
        m_reverseFlag = 0
    End If

    ''''<NOTICE> outBitWave, its Element(0) is always LSBbit value
    m_tmpWave = m_doubleWave.Select(m_startbit, 1, m_bitwidth).Copy
    If (m_reverseFlag = 0) Then
        outBitWave = m_tmpWave.Copy
    Else
        ''''do reverse
        Call eFuse_reverseBitWave(m_tmpWave, outBitWave)
    End If
    outbitSum = outBitWave.CalcSum

    m_tmpArr = m_readCateWave.Data
    outDec = m_tmpArr(idx)
    
    If (resolution = 0#) Then resolution = 1#

    If (calc_mode = 0) Then
        ''''ids: resolution<>0
        ''''decimal: resolution=1
        outVal = outDec * resolution
    ElseIf (calc_mode = 1) Then
        ''''vddbin and safe voltage
        outVal = gDD_BaseVoltage + (outDec * resolution)
    ElseIf (calc_mode = 2) Then
        ''''bincut, only limit is variant per dice
        outVal = gDD_BaseVoltage + (outDec * resolution)
    ElseIf (calc_mode = 3) Then
        ''''base and safe voltage
        outVal = (outDec + 1) * resolution
    End If

End Function

Public Function eFuse_compare_MarginRead_DoubleBitWave(ByVal FuseType As Long, ByVal InWave As DSPWave, _
                                                           ByVal Reverse As Long, ByRef cmpResult As Long) As Long
On Error Resume Next
    ''''------------------------------------------------------------------------------
    ''''Definition of the Variables
    ''''------------------------------------------------------------------------------
    ''''     inWave: eFuse DAP/JTAG Capture Data Wave
    ''''    reverse: 0 (normal), 1 (reverse)
    ''''  cmpResult: 0 means Equal (Read=DAP), >0 means Not Equal (Read<>DAP)
    ''''------------------------------------------------------------------------------
    
    Dim m_size As Long
    Dim sgSum As Long
    Dim dbSum As Long
    Dim m_cmpResWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_readBitWave As New DSPWave

    If (Reverse = 1) Then
        Call eFuse_reverseBitWave(InWave, m_tmpWave)
    Else
        m_tmpWave = InWave.Copy
    End If

    If (FuseType = 1) Then
        m_readBitWave = gDW_ECID_Read_DoubleBitWave.Copy
    ElseIf (FuseType = 2) Then
        m_readBitWave = gDW_CFG_Read_DoubleBitWave.Copy
    ElseIf (FuseType = 3) Then
    ElseIf (FuseType = 4) Then
    ElseIf (FuseType = 5) Then
    ElseIf (FuseType = 6) Then
    ElseIf (FuseType = 7) Then
    ElseIf (FuseType = 8) Then
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If

    m_cmpResWave = m_readBitWave.LogicalCompare(NotEqualTo, m_tmpWave)
    
    ''''Here LogicalCompare(NotEqualTo) return 1, otherwise return 0 (equal) for each element
    ''''cmpResult means how many bits are NOt equal
    ''''if (cmpResult=0), means Read==inWave
    cmpResult = m_cmpResWave.CalcSum

End Function

Public Function eFuse_compare_MarginRead_SingleBitWave(ByVal FuseType As Long, ByVal InWave As DSPWave, _
                                                       ByVal Reverse As Long, ByRef cmpResult As Long) As Long
On Error Resume Next
                                                       
    ''''------------------------------------------------------------------------------
    ''''Definition of the Variables
    ''''------------------------------------------------------------------------------
    ''''     inWave: eFuse DAP/JTAG Capture Data Wave
    ''''    reverse: 0 (normal), 1 (reverse)
    ''''  cmpResult: 0 means Equal (Read=DAP), >0 means Not Equal (Read<>DAP)
    ''''------------------------------------------------------------------------------
    
    Dim m_size As Long
    Dim sgSum As Long
    Dim dbSum As Long
    Dim m_cmpResWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_readBitWave As New DSPWave

    If (Reverse = 1) Then
        Call eFuse_reverseBitWave(InWave, m_tmpWave)
    Else
        m_tmpWave = InWave.Copy
    End If

    If (FuseType = 1) Then
        m_readBitWave = gDW_ECID_Read_SingleBitWave.Copy
    ElseIf (FuseType = 2) Then
        m_readBitWave = gDW_CFG_Read_SingleBitWave.Copy
    ElseIf (FuseType = 3) Then
    ElseIf (FuseType = 4) Then
    ElseIf (FuseType = 5) Then
    ElseIf (FuseType = 6) Then
    ElseIf (FuseType = 7) Then
    ElseIf (FuseType = 8) Then
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If

    m_cmpResWave = m_readBitWave.LogicalCompare(NotEqualTo, m_tmpWave)
    
    ''''Here LogicalCompare(NotEqualTo) return 1, otherwise return 0 (equal) for each element
    ''''cmpResult means how many bits are NOt equal
    ''''if (cmpResult=0), means Read==inWave
    cmpResult = m_cmpResWave.CalcSum

End Function

Private Function eFuse_CRC32_ComputeCRCforBit_DSP(ByVal calcbit As Long, ByRef crcBitWave As DSPWave) As Long
On Error Resume Next

    Dim m_invBit As Long
    Dim m_crcArr() As Long
    Dim m_crcWave As New DSPWave

    m_crcWave = crcBitWave.Copy
    m_crcArr = m_crcWave.Data
    
    m_invBit = calcbit Xor m_crcArr(31)
    m_crcArr(31) = m_crcArr(30)
    m_crcArr(30) = m_crcArr(29)
    m_crcArr(29) = m_crcArr(28)
    m_crcArr(28) = m_crcArr(27)
    m_crcArr(27) = m_crcArr(26)
    m_crcArr(26) = m_crcArr(25) Xor m_invBit
    m_crcArr(25) = m_crcArr(24)
    m_crcArr(24) = m_crcArr(23)
    m_crcArr(23) = m_crcArr(22) Xor m_invBit
    m_crcArr(22) = m_crcArr(21) Xor m_invBit
    m_crcArr(21) = m_crcArr(20)
    m_crcArr(20) = m_crcArr(19)
    m_crcArr(19) = m_crcArr(18)
    m_crcArr(18) = m_crcArr(17)
    m_crcArr(17) = m_crcArr(16)
    m_crcArr(16) = m_crcArr(15) Xor m_invBit
    m_crcArr(15) = m_crcArr(14)
    m_crcArr(14) = m_crcArr(13)
    m_crcArr(13) = m_crcArr(12)
    m_crcArr(12) = m_crcArr(11) Xor m_invBit
    m_crcArr(11) = m_crcArr(10) Xor m_invBit
    m_crcArr(10) = m_crcArr(9) Xor m_invBit
    m_crcArr(9) = m_crcArr(8)
    m_crcArr(8) = m_crcArr(7) Xor m_invBit
    m_crcArr(7) = m_crcArr(6) Xor m_invBit
    m_crcArr(6) = m_crcArr(5)
    m_crcArr(5) = m_crcArr(4) Xor m_invBit
    m_crcArr(4) = m_crcArr(3) Xor m_invBit
    m_crcArr(3) = m_crcArr(2)
    m_crcArr(2) = m_crcArr(1) Xor m_invBit
    m_crcArr(1) = m_crcArr(0) Xor m_invBit
    m_crcArr(0) = m_invBit

    m_crcWave.Data = m_crcArr
    crcBitWave = m_crcWave.Copy

''''--------------------------------------------
'''' Reference CRC_ComputeCRCforBit()
''''--------------------------------------------
''''    inv = bit Xor CRC(31)
''''    CRC(31) = CRC(30)
''''    CRC(30) = CRC(29)
''''    CRC(29) = CRC(28)
''''    CRC(28) = CRC(27)
''''    CRC(27) = CRC(26)
''''    CRC(26) = CRC(25) Xor inv
''''    CRC(25) = CRC(24)
''''    CRC(24) = CRC(23)
''''    CRC(23) = CRC(22) Xor inv
''''    CRC(22) = CRC(21) Xor inv
''''    CRC(21) = CRC(20)
''''    CRC(20) = CRC(19)
''''    CRC(19) = CRC(18)
''''    CRC(18) = CRC(17)
''''    CRC(17) = CRC(16)
''''    CRC(16) = CRC(15) Xor inv
''''    CRC(15) = CRC(14)
''''    CRC(14) = CRC(13)
''''    CRC(13) = CRC(12)
''''    CRC(12) = CRC(11) Xor inv
''''    CRC(11) = CRC(10) Xor inv
''''    CRC(10) = CRC(9) Xor inv
''''    CRC(9) = CRC(8)
''''    CRC(8) = CRC(7) Xor inv
''''    CRC(7) = CRC(6) Xor inv
''''    CRC(6) = CRC(5)
''''    CRC(5) = CRC(4) Xor inv
''''    CRC(4) = CRC(3) Xor inv
''''    CRC(3) = CRC(2)
''''    CRC(2) = CRC(1) Xor inv
''''    CRC(1) = CRC(0) Xor inv
''''    CRC(0) = inv
''''--------------------------------------------
End Function

Private Function eFuse_CRC16_ComputeCRCforBit_DSP(ByVal calcbit As Long, ByRef crcBitWave As DSPWave) As Long
On Error Resume Next

    Dim m_invBit As Long
    Dim m_crcArr() As Long
    Dim m_crcWave As New DSPWave

    m_crcWave = crcBitWave.Copy
    m_crcArr = m_crcWave.Data
    
    m_invBit = calcbit Xor m_crcArr(15)
    m_crcArr(15) = m_crcArr(14)
    m_crcArr(14) = m_crcArr(13)
    m_crcArr(13) = m_crcArr(12) Xor m_invBit
    m_crcArr(12) = m_crcArr(11) Xor m_invBit
    m_crcArr(11) = m_crcArr(10) Xor m_invBit
    m_crcArr(10) = m_crcArr(9) Xor m_invBit
    m_crcArr(9) = m_crcArr(8)
    m_crcArr(8) = m_crcArr(7) Xor m_invBit
    m_crcArr(7) = m_crcArr(6)
    m_crcArr(6) = m_crcArr(5) Xor m_invBit
    m_crcArr(5) = m_crcArr(4) Xor m_invBit
    m_crcArr(4) = m_crcArr(3)
    m_crcArr(3) = m_crcArr(2)
    m_crcArr(2) = m_crcArr(1) Xor m_invBit
    m_crcArr(1) = m_crcArr(0)
    m_crcArr(0) = m_invBit

    m_crcWave.Data = m_crcArr
    crcBitWave = m_crcWave.Copy
    
''''--------------------------------------------
'''' Reference CRC16_ComputeCRCforBit()
''''--------------------------------------------
''''    Dim inv As Byte
''''    inv = bit Xor CRC(15)
''''    CRC(15) = CRC(14)
''''    CRC(14) = CRC(13)
''''    CRC(13) = CRC(12) Xor inv
''''    CRC(12) = CRC(11) Xor inv
''''    CRC(11) = CRC(10) Xor inv
''''    CRC(10) = CRC(9) Xor inv
''''    CRC(9) = CRC(8)
''''    CRC(8) = CRC(7) Xor inv
''''    CRC(7) = CRC(6)
''''    CRC(6) = CRC(5) Xor inv
''''    CRC(5) = CRC(4) Xor inv
''''    CRC(4) = CRC(3)
''''    CRC(3) = CRC(2)
''''    CRC(2) = CRC(1) Xor inv
''''    CRC(1) = CRC(0)
''''    CRC(0) = inv
''''--------------------------------------------
End Function

Public Function eFuse_CRC_Calculation_DSP(ByVal FuseType As Long, ByVal BitWidth As Long, ByVal InWave As DSPWave, ByRef crcBitWave As DSPWave, _
                                          ByRef bitForcrcCalcWave As DSPWave) As Long
On Error Resume Next

    Dim j As Long
    Dim m_size As Long
    Dim m_cnt As Long
    Dim m_calcBitFlagArr() As Long
    Dim m_tmpArr() As Long
    Dim m_tmpWave As New DSPWave
    Dim m_crcWave As New DSPWave
    Dim m_bitForcrcCalcArr() As Long

    ''''---------------------------------------------------------------------------
    '''' inWave: is a doubleBitWave (Pgm OR Read)
    ''''---------------------------------------------------------------------------
    ''''Methodolog
    ''''---------------------------------------------------------------------------
    ''''using calcBits Array to control the CRC calculated bits
    ''''m_calcBitFlagArr(), CRC calcBits Array, =1 means CRC calculated bit,
    ''''                                        =0 means CRC ignore bit.
    ''''---------------------------------------------------------------------------
    
''''''''-----------------------------------
''''''''From Enum eFuseBlockType
''''    eFuse_ECID = 1
''''    eFuse_CFG = 2
''''    eFuse_UID = 3
''''    eFuse_SEN = 4
''''    eFuse_MON = 5
''''    eFuse_UDR = 6
''''    eFuse_UDRE = 7
''''    eFuse_UDRP = 8
''''    eFuse_CMP = 9
''''    eFuse_CMPE = 10
''''    eFuse_CMPP = 11
''''    eFuse_Block_Unknown = 999
''''''''-----------------------------------

    m_crcWave.CreateConstant 0, BitWidth, DspLong
    m_tmpWave = InWave.Copy
    m_size = m_tmpWave.SampleSize
    m_tmpArr = m_tmpWave.Data
    
    If (FuseType = 1) Then
        m_calcBitFlagArr = gDW_ECID_CRC_calcBitsWave.Data
    ElseIf (FuseType = 2) Then
        m_calcBitFlagArr = gDW_CFG_CRC_calcBitsWave.Data
    ElseIf (FuseType = 3) Then
        m_calcBitFlagArr = gDW_UID_CRC_calcBitsWave.Data
    ElseIf (FuseType = 4) Then
        m_calcBitFlagArr = gDW_SEN_CRC_calcBitsWave.Data
    ElseIf (FuseType = 5) Then
        m_calcBitFlagArr = gDW_MON_CRC_calcBitsWave.Data
    Else
    End If

    m_cnt = 0
    ReDim m_bitForcrcCalcArr(UBound(m_calcBitFlagArr))
    
    If (BitWidth = 16) Then
        For j = (m_size - 1) To 0 Step -1
            If (m_calcBitFlagArr(j) = 1) Then
                Call eFuse_CRC16_ComputeCRCforBit_DSP(m_tmpArr(j), m_crcWave)
                m_bitForcrcCalcArr(m_cnt) = m_tmpArr(j)
                m_cnt = m_cnt + 1
            End If
        Next j
    ElseIf (BitWidth = 32) Then
        For j = (m_size - 1) To 0 Step -1
            If (m_calcBitFlagArr(j) = 1) Then
                Call eFuse_CRC32_ComputeCRCforBit_DSP(m_tmpArr(j), m_crcWave)
                m_bitForcrcCalcArr(m_cnt) = m_tmpArr(j)
                m_cnt = m_cnt + 1
            End If
        Next j
    End If

    ''''return CRC result
    crcBitWave = m_crcWave.Copy
    
    ReDim Preserve m_bitForcrcCalcArr(m_cnt - 1)
    bitForcrcCalcWave.Data = m_bitForcrcCalcArr

End Function

''''201812XX New, it's for the real case.
Public Function eFuse_updatePgmWave_CRCbits(ByVal FuseType As Long, ByVal BitWidth As Long, ByVal indexWave As DSPWave) As Long
On Error Resume Next

    Dim m_inWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_tmpWave2 As New DSPWave
    Dim m_defaultBitWave As New DSPWave
    Dim m_effbitFlagWave As New DSPWave
    Dim m_readBitWave As New DSPWave
    Dim m_crcBitWave As New DSPWave
    Dim m_bitForcrcCalcWave As New DSPWave

''''''''-----------------------------------
''''''''From Enum eFuseBlockType
''''    eFuse_ECID = 1
''''    eFuse_CFG = 2
''''    eFuse_UID = 3
''''    eFuse_SEN = 4
''''    eFuse_MON = 5
''''    eFuse_UDR = 6
''''    eFuse_UDRE = 7
''''    eFuse_UDRP = 8
''''    eFuse_CMP = 9
''''    eFuse_CMPE = 10
''''    eFuse_CMPP = 11
''''    eFuse_Block_Unknown = 999
''''''''-----------------------------------

    If (FuseType = 1) Then
        m_readBitWave = gDW_ECID_Read_DoubleBitWave.Copy
        m_defaultBitWave = gDW_ECID_allDefaultBitWave.Copy
        m_effbitFlagWave = gDW_ECID_Stage_BitFlag.Copy
    ElseIf (FuseType = 2) Then
        m_readBitWave = gDW_CFG_Read_DoubleBitWave.Copy
        m_defaultBitWave = gDW_CFG_allDefaultBitWave.Copy
        m_effbitFlagWave = gDW_CFG_Stage_BitFlag.Copy
    ElseIf (FuseType = 3) Then
        m_readBitWave = gDW_UID_Read_DoubleBitWave.Copy
        m_defaultBitWave = gDW_UID_allDefaultBitWave.Copy
        m_effbitFlagWave = gDW_UID_Stage_BitFlag.Copy
    ElseIf (FuseType = 4) Then
        m_readBitWave = gDW_SEN_Read_DoubleBitWave.Copy
        m_defaultBitWave = gDW_SEN_allDefaultBitWave.Copy
        m_effbitFlagWave = gDW_SEN_Stage_BitFlag.Copy
    ElseIf (FuseType = 5) Then
        m_readBitWave = gDW_MON_Read_DoubleBitWave.Copy
        m_defaultBitWave = gDW_MON_allDefaultBitWave.Copy
        m_effbitFlagWave = gDW_MON_Stage_BitFlag.Copy
    Else
    End If
    
    ''''<MUST for simulation> is happened on simulation case (blank=False 1st time)
    If (m_readBitWave.SampleSize = 0) Then
        m_readBitWave.CreateConstant 0, gDW_Pgm_RawBitWave.SampleSize, DspLong
    End If
    
    m_tmpWave = gDW_Pgm_RawBitWave.Copy ''''Here it only includes "Real" pgm bits

    ''''<MUST> including defaultBits for the CRC calculation
    ''''<BeCareful> m_effbitFlagWave, could be different online/offline mode. [check it later]20181221
    m_tmpWave2 = m_tmpWave.BitwiseOr(m_defaultBitWave).bitwiseand(m_effbitFlagWave).Copy
    m_inWave = m_tmpWave2.BitwiseOr(m_readBitWave).Copy
    Call eFuse_CRC_Calculation_DSP(FuseType, BitWidth, m_inWave, m_crcBitWave, m_bitForcrcCalcWave)
    If (FuseType = 1) Then
        Dim outWave As New DSPWave
        Call eFuse_reverseBitWave(m_crcBitWave, outWave)
        Call m_tmpWave.ReplaceElements(indexWave, outWave)
    Else
        Call m_tmpWave.ReplaceElements(indexWave, m_crcBitWave)
    End If
    
    gDW_Pgm_RawBitWave = m_tmpWave.Copy
    
    gDW_Pgm_BitWaveForCRCCalc = m_bitForcrcCalcWave.Copy

End Function

''''201812XX New, it's for the simulation mode.
Public Function eFuse_updatePgmWave_CRCbits_Simulation(ByVal FuseType As Long, ByVal BitWidth As Long, ByVal indexWave As DSPWave) As Long
On Error Resume Next

    Dim m_inWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_tmpWave2 As New DSPWave
    Dim m_defaultBitWave As New DSPWave
    Dim m_effbitFlagWave As New DSPWave
    Dim m_readBitWave As New DSPWave
    Dim m_crcBitWave As New DSPWave
    Dim m_bitForcrcCalcWave As New DSPWave

''''''''-----------------------------------
''''''''From Enum eFuseBlockType
''''    eFuse_ECID = 1
''''    eFuse_CFG = 2
''''    eFuse_UID = 3
''''    eFuse_SEN = 4
''''    eFuse_MON = 5
''''    eFuse_UDR = 6
''''    eFuse_UDRE = 7
''''    eFuse_UDRP = 8
''''    eFuse_CMP = 9
''''    eFuse_CMPE = 10
''''    eFuse_CMPP = 11
''''    eFuse_Block_Unknown = 999
''''''''-----------------------------------

    If (FuseType = 1) Then
        m_readBitWave = gDW_ECID_Read_DoubleBitWave.Copy
        m_defaultBitWave = gDW_ECID_allDefaultBitWave.Copy
        m_effbitFlagWave = gDW_ECID_StageLEQJob_BitFlag.Copy ''gDW_ECID_Stage_BitFlag.Copy
    ElseIf (FuseType = 2) Then
        m_readBitWave = gDW_CFG_Read_DoubleBitWave.Copy
        m_defaultBitWave = gDW_CFG_allDefaultBitWave.Copy
        m_effbitFlagWave = gDW_CFG_StageLEQJob_BitFlag.Copy ''gDW_CFG_Stage_BitFlag.Copy
    ElseIf (FuseType = 3) Then
    ElseIf (FuseType = 4) Then
    ElseIf (FuseType = 5) Then
        m_readBitWave = gDW_MON_Read_DoubleBitWave.Copy
        m_defaultBitWave = gDW_MON_allDefaultBitWave.Copy
        m_effbitFlagWave = gDW_MON_StageLEQJob_BitFlag.Copy ''gDW_CFG_Stage_BitFlag.Copy
    Else
    End If
    
    ''''<MUST for simulation> is happened on simulation case (blank=False 1st time)
    If (m_readBitWave.SampleSize = 0) Then
        m_readBitWave.CreateConstant 0, gDW_Pgm_RawBitWave.SampleSize, DspLong
    End If
    
    m_tmpWave = gDW_Pgm_RawBitWave.Copy ''''Here it only includes "Real" pgm bits

    ''''<MUST> including defaultBits for the CRC calculation
    ''''<BeCareful> m_effbitFlagWave, could be different online/offline mode. [check it later]20181221
    m_tmpWave2 = m_tmpWave.BitwiseOr(m_defaultBitWave).bitwiseand(m_effbitFlagWave).Copy
    m_inWave = m_tmpWave2.BitwiseOr(m_readBitWave).Copy
    Call eFuse_CRC_Calculation_DSP(FuseType, BitWidth, m_inWave, m_crcBitWave, m_bitForcrcCalcWave)

    Call m_tmpWave.ReplaceElements(indexWave, m_crcBitWave)
    gDW_Pgm_RawBitWave = m_tmpWave.Copy
    
    gDW_Pgm_BitWaveForCRCCalc = m_bitForcrcCalcWave.Copy

End Function

Public Function eFuse_Read_to_calc_CRCWave(ByVal FuseType As Long, ByVal BitWidth As Long, ByRef crcBitWave As DSPWave) As Long
On Error Resume Next

    Dim m_inWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_readBitWave As New DSPWave
    Dim m_crcBitWave As New DSPWave
    Dim m_bitForcrcCalcWave As New DSPWave

''''''''-----------------------------------
''''''''From Enum eFuseBlockType
''''    eFuse_ECID = 1
''''    eFuse_CFG = 2
''''    eFuse_UID = 3
''''    eFuse_SEN = 4
''''    eFuse_MON = 5
''''    eFuse_UDR = 6
''''    eFuse_UDRE = 7
''''    eFuse_UDRP = 8
''''    eFuse_CMP = 9
''''    eFuse_CMPE = 10
''''    eFuse_CMPP = 11
''''    eFuse_Block_Unknown = 999
''''''''-----------------------------------

    If (FuseType = 1) Then
        m_readBitWave = gDW_ECID_Read_DoubleBitWave.Copy
    ElseIf (FuseType = 2) Then
        m_readBitWave = gDW_CFG_Read_DoubleBitWave.Copy
    ElseIf (FuseType = 3) Then
        m_readBitWave = gDW_UID_Read_DoubleBitWave.Copy
    ElseIf (FuseType = 4) Then
        m_readBitWave = gDW_SEN_Read_DoubleBitWave.Copy
    ElseIf (FuseType = 5) Then
        m_readBitWave = gDW_MON_Read_DoubleBitWave.Copy
    Else
    End If
    
    Call eFuse_CRC_Calculation_DSP(FuseType, BitWidth, m_readBitWave, m_crcBitWave, m_bitForcrcCalcWave)
    
    crcBitWave = m_crcBitWave.Copy
    gDW_Read_BitWaveForCRCCalc = m_bitForcrcCalcWave.Copy

End Function
''''''''''''''''the below is a trial, but NOT working
''''            If (True) Then
''''                ''''the below is a trial
''''                Dim m_Result As Boolean
''''                Dim m_cmpWave As New DSPWave
''''                Dim m_tmpWave2 As New DSPWave
''''                m_bitwidth = 112
''''                m_tmpWave2.CreateRandom 0, 1, m_bitwidth, 1, DspLong
''''                m_decimal = m_tmpWave2.ConvertStreamTo(tldspParallel, m_bitwidth, 0, Bit0IsLsb).ConvertDataTypeTo(DspDouble).Element(0)
''''
''''                Dim m_quo As Double
''''                Dim m_rest As Double
''''                Dim m_bw As Long
''''                Dim m_cnt As Long
''''                Dim m_mod As Long
''''                Dim m_divider As Double
''''                m_cnt = Floor(m_bitwidth / 32)
''''                m_mod = m_bitwidth Mod 32
''''                If (m_mod = 0) Then m_cnt = m_cnt - 1
''''
''''                Dim m_dbl As Double
''''                'm_decimal = 256
''''                m_cnt = 0
''''                Do
''''                    m_rest = CDbl(m_decimal / 2) - Floor(m_decimal / 2)
''''                    If (m_rest = 0) Then
''''                        Debug.Print "bit_" & m_cnt & " = " & 0
''''                    Else
''''                        Debug.Print "bit_" & m_cnt & " = " & 1
''''                    End If
''''                    m_decimal = Floor(m_decimal / 2)
''''                    m_cnt = m_cnt + 1
''''                Loop Until (m_decimal <= 0)
                
''''                For j = 1 To m_cnt
''''                    Debug.Print m_decimal
''''
''''                    m_bw = m_bitwidth - (32 * 1)
''''                    m_divider = CDbl(2 ^ m_bw)
''''                    m_quo = m_decimal / m_divider
''''                    m_rest = m_decimal - (Floor(m_quo) * m_divider)
''''                    Debug.Print j & " , " & m_decimal & " / " & m_divider & " = " & m_quo & " ,rest=" & m_rest & ", bw=" & m_bw
''''
''''                    m_bw = m_bitwidth - (32 * (j + 1))
''''                    m_divider = CDbl(2 ^ m_bw)
''''                    m_quo = m_rest / m_divider
''''                    m_rest = m_rest - (Floor(m_quo) * m_divider)
''''                    Debug.Print (j + 1) & " , " & m_rest & " / " & m_divider & " = " & m_quo & " ,rest=" & m_rest & ", bw=" & m_bw
''''
''''                Next j
''''            End If
''''-------------------------------------------------------------------------
'''''Not Working
''''Public Function eFuse_DspWaveArr_Copy(ByVal inWave0 As DSPWave, ByVal inWave1 As DSPWave, ByVal inWave2 As DSPWave, ByVal arrSize As Long, ByRef outWave As DSPWave) As Long
''''
''''    ReDim outWave(arrSize - 1)
''''
''''    outWave(0) = inWave0.ConvertDataTypeTo(DspLong).Copy
''''    outWave(1) = inWave1.ConvertDataTypeTo(DspLong).Copy
''''    outWave(2) = inWave2.ConvertDataTypeTo(DspLong).Copy
''''
''''End Function
''''-------------------------------------------------------------------------

Public Function eFuse_Gen_PgmBitSrcWave_OnlyRV(ByVal FuseType As Long, ByVal bitFlag_mode As Long, _
                                        ByVal RV_CNT As Long, _
                                        ByRef outSrcWave As DSPWave, ByRef result As Long) As Long
On Error Resume Next

    '''' bitFlag_mode: 0=early_stage, 1=normal stage to decide which bitFlagWave(Stage_early or Stage)
    
    Dim i As Long, j As Long, k As Long
    Dim k1 As Long, k2 As Long
    Dim m_tmpValue As Long
    Dim m_sgbits As Long
    
    Dim expandWidth As Long
    Dim BitsPerRow As Long
    Dim ReadCycles As Long
    Dim BitsPerCycle As Long
    Dim BitsPerBlock As Long
    
    Dim m_tmpWave1 As New DSPWave
    Dim m_tmpWave2 As New DSPWave
    Dim m_pgmrawBitWave As New DSPWave
    Dim m_defaultBitWave As New DSPWave
    Dim m_effbitFlagWave As New DSPWave
    Dim m_singleWave As New DSPWave
    Dim m_doubleWave As New DSPWave
    Dim m_RVWave As New DSPWave
    Dim m_RVSampleSize As Long
    
    Dim m_size As Long
    Dim m_tmpArr() As Long
    Dim m_dbbitArr() As Long
    Dim m_sgBitArr() As Long
    Dim m_outSrcBitArr() As Long
    
    BitsPerRow = gDL_BitsPerRow
    ReadCycles = gDL_ReadCycles
    BitsPerCycle = gDL_BitsPerCycle
    BitsPerBlock = gDL_BitsPerBlock
    expandWidth = gDL_DigSrcRepeatN

''''''''-----------------------------------
''''''''From Enum eFuseBlockType
''''    eFuse_ECID = 1
''''    eFuse_CFG = 2
''''    eFuse_UID = 3
''''    eFuse_SEN = 4
''''    eFuse_MON = 5
''''    eFuse_UDR = 6
''''    eFuse_UDRE = 7
''''    eFuse_UDRP = 8
''''    eFuse_CMP = 9
''''    eFuse_CMPE = 10
''''    eFuse_CMPP = 11
''''    eFuse_Block_Unknown = 999
''''''''-----------------------------------
    
    If (FuseType = 1) Then
        'm_defaultBitWave = gDW_ECID_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_ECID_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_ECID_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 2) Then
        'm_defaultBitWave = gDW_CFG_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_CFG_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_CFG_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 3) Then
       ' m_defaultBitWave = gDW_UID_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_UID_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_UID_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 4) Then
        'm_defaultBitWave = gDW_SEN_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_SEN_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_SEN_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 5) Then
        'm_defaultBitWave = gDW_MON_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_MON_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_MON_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 6) Then
        'm_defaultBitWave = gDW_UDR_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_UDR_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_UDR_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 7) Then
        'm_defaultBitWave = gDW_UDRE_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_UDRE_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_UDRE_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 8) Then
        'm_defaultBitWave = gDW_UDRP_allDefaultBitWave.Copy
        If (bitFlag_mode = 0) Then
            m_effbitFlagWave = gDW_UDRP_Stage_Early_BitFlag.Copy
        ElseIf (bitFlag_mode = 1) Then
            m_effbitFlagWave = gDW_UDRP_Stage_BitFlag.Copy
        End If
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If

    If (bitFlag_mode > 1) Then
        ''''use all bits
        m_effbitFlagWave.CreateConstant 1, BitsPerBlock, DspLong
    End If

   m_tmpWave1 = gDW_Pgm_RawBitWave.Copy

    ''''combine m_pgmrawBitWave with "OR" the m_defaultBitWave
    'm_tmpWave1 = m_pgmrawBitWave.BitwiseOr(m_defaultBitWave).Copy
    
    ''''gen effective Wave with "AND" the effbitFlagWave
    m_doubleWave = m_tmpWave1.bitwiseand(m_effbitFlagWave).Copy

    m_sgbits = BitsPerCycle * ReadCycles
    m_RVSampleSize = RV_CNT * 32
    
    m_singleWave.CreateConstant 0, m_sgbits, DspLong
    m_RVWave.CreateConstant 0, m_RVSampleSize, DspLong
    
    If (gDL_eFuse_Orientation = 0) Then ''''Up2Down
        m_singleWave = m_doubleWave.repeat(2).Copy

    ElseIf (gDL_eFuse_Orientation = 1) Then ''''SingleUp, eFuse_1_Bit
        ''''doubleBitWave is equal to singleBitWave
        m_singleWave = m_doubleWave.Copy

    ElseIf (gDL_eFuse_Orientation = 2) Then ''''Right2Left, eFuse_2_Bit
        ''''-------------------------------------------------------------------------------
        ''''New Method
        ''''-------------------------------------------------------------------------------
        If (True) Then
            Dim m_EveryRowWave As New DSPWave
            Dim m_TmpRowWave As New DSPWave
            Dim m_SegFlag() As Long
            m_SegFlag = gDW_CFG_SegFlag.Data
        
            k = 0 ''''MUST
            For i = 0 To gDW_CFG_SegFlag.SampleSize - 1
                If (m_SegFlag(i) = 1) Then
                    j = k * 32
                    m_TmpRowWave = m_doubleWave.Select(16 * i, 1, 16)
                    m_EveryRowWave = m_TmpRowWave.Concatenate(m_TmpRowWave).Copy
                    m_RVWave.Select(j, 1, 32).Replace m_EveryRowWave
                    m_EveryRowWave.CreateConstant 0, 32, DspLong
                    k = k + 1
                End If
            Next i
        
        
       ' Else
            m_dbbitArr = m_doubleWave.Data
            m_sgBitArr = m_singleWave.Data
    
            k = 0 ''''must be here
            For i = 0 To ReadCycles - 1     'EX: EcidReadCycle - 1   ''0...15(ECID)
                For j = 0 To BitsPerRow - 1 'EX: EcidBitsPerRow - 1  ''0...15(ECID)
                    ''''k1: Right block
                    ''''k2:  Left block
                    k1 = (i * BitsPerCycle) + j ''<Important> Must use BitsPerCycle here
                    k2 = (i * BitsPerCycle) + BitsPerRow + j
                    ''m_tmpValue = m_doubleWave.Element(k)
                    ''m_singleWave.Element(k1) = m_tmpValue
                    ''m_singleWave.Element(k2) = m_tmpValue
                    
                    ''''save TT here by using dataArr
                    m_tmpValue = m_dbbitArr(k)
                    m_sgBitArr(k1) = m_tmpValue
                    m_sgBitArr(k2) = m_tmpValue
                    k = k + 1
                Next j
            Next i
            m_singleWave.Data = m_sgBitArr ''''save TT
        End If
        ''''-------------------------------------------------------------------------------

    ''''Below is reserved for the future, there is NO any definition at present.
    ElseIf (gDL_eFuse_Orientation = 3) Then
    ElseIf (gDL_eFuse_Orientation = 4) Then
    ElseIf (gDL_eFuse_Orientation = 5) Then
    End If

    ''''Expand the outWave to Source Wave
    m_size = m_sgbits * expandWidth
'    outSrcWave.CreateConstant 0, m_size, DspLong
    outSrcWave.CreateConstant 0, m_RVSampleSize, DspLong
    ReDim m_outSrcBitArr(m_size - 1)
    
    If (gDL_eFuse_Orientation <> 2) Then m_sgBitArr = m_singleWave.Data
    
'    k = 0
'    For i = 0 To m_sgbits - 1
'        For j = 0 To expandWidth - 1
'            k = i * expandWidth + j
'            ''outSrcWave.Element(k) = m_singleWave.Element(i) ''''waste TT
'            m_outSrcBitArr(k) = m_sgBitArr(i) ''''to save TT
'        Next j
'    Next i
'    outSrcWave.Data = m_outSrcBitArr ''''to save TT
    'outSrcWave.Data = m_singleWave
    outSrcWave = m_RVWave.Copy
    ''''---------------------------------------------------------------
    ''''Update to global DSP variables
    ''''---------------------------------------------------------------
    If (FuseType = 1) Then
        gDW_ECID_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_ECID_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 2) Then
        gDW_CFG_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_CFG_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 3) Then
        gDW_UID_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_UID_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 4) Then
        gDW_SEN_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_SEN_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 5) Then
        gDW_MON_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_MON_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 6) Then
        gDW_UDR_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_UDR_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 7) Then
        gDW_UDRE_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_UDRE_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 8) Then
        gDW_UDRP_Pgm_SingleBitWave = m_singleWave.Copy
        gDW_UDRP_Pgm_DoubleBitWave = m_doubleWave.Copy
    ElseIf (FuseType = 9) Then
    ElseIf (FuseType = 10) Then
    ElseIf (FuseType = 11) Then
    Else
    End If
    ''''---------------------------------------------------------------

    result = 1

End Function

