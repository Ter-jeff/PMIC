Attribute VB_Name = "DSP_OTP_LIB"
'T-AutoGen-Version : 1.3.0.1
'ProjectName_A1_TestPlan_20220226.xlsx
'ProjectName_A0_otp_AVA.otp
'ProjectName_A0_OTP_register_map.yaml
'ProjectName_A0_Pattern_List_Ext_20190823.csv
'ProjectName_A0_scgh_file#1_20200207.xlsx
'ProjectName_A0_VBTPOP_Gen_tool_MP10P_BuckSW_UVI80_DiffMeter_20200430.xlsm
Option Explicit

' This module should be used only for DSP Procedure code.  Functions in this
' module will be available to be called to perform DSP in all DSP modes.
' Additional modules may be added as needed (all starting with "DSP_").
'
' The required signature for a DSP Procedure is:
'
' Public Function FuncName(<arglist>) as Long
'   where <arglist> is any list of arguments supported by DSP code.
'
' See online help for supported types and other restrictions.
'**********************************************************************************************

'___Concatenate the addr and data dspwave
Public Function otp_get_pgm_AddrDataWave(ByVal lAddrVal As Long, ByVal lAddrIdx As Long, ByRef r_wAddrAndData As DSPWave) As Long
    On Error Resume Next
    Dim wAddr As New DSPWave
    Dim wAddrBit As New DSPWave
    Dim wDataBit As New DSPWave
    
    wAddr.CreateConstant lAddrVal, 1, DspLong
    wAddrBit = wAddr.ConvertStreamTo(tldspSerial, gD_slOTP_ADDR_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
    wDataBit = gD_wPGMData.Select(lAddrIdx * gD_slOTP_REGDATA_BW, 1, gD_slOTP_REGDATA_BW).Copy
    r_wAddrAndData = wAddrBit.Concatenate(wDataBit)
    
Exit Function
End Function

''''20190725, update method for the FTProg to A0 except Lockbit to prevent burn '1' again
'___Concatenate the addr and data dspwave
Public Function otp_get_pgm_AddrDataWave_maskECID(ByVal lAddrVal As Long, ByVal lAddrIdx As Long, ByVal sbOtpedECID As Boolean, ByVal wPgmWaveMaskECID As DSPWave, ByRef r_wAddrAndData As DSPWave) As Long
    On Error Resume Next
    Dim wAddr As New DSPWave
    Dim wAddrBit As New DSPWave
    Dim wDataBit As New DSPWave

    wAddr.CreateConstant lAddrVal, 1, DspLong
    wAddrBit = wAddr.ConvertStreamTo(tldspSerial, gD_slOTP_ADDR_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
    
    '''' sbOtpedECID = True  , means ECID not burned => g_sbOtpedECID = True
    '''' sbOtpedECID = False , means ECID burned     => g_sbOtpedECID = False
    If sbOtpedECID Then
        ''''because ECID has burned, so we set all ECID bits as "0"
        wDataBit = wPgmWaveMaskECID.Select(lAddrIdx * gD_slOTP_REGDATA_BW, 1, gD_slOTP_REGDATA_BW).Copy
        r_wAddrAndData = wAddrBit.Concatenate(wDataBit)
    Else
        wDataBit = gD_wPGMData.Select(lAddrIdx * gD_slOTP_REGDATA_BW, 1, gD_slOTP_REGDATA_BW).Copy
        r_wAddrAndData = wAddrBit.Concatenate(wDataBit)
    End If
Exit Function
End Function

''''New 20190329 OTP-DSP
'___Return a decimal read value decided by the section of gD_wReadData
Public Function otp_get_read_DataWave(ByVal wDataBit As DSPWave, ByVal lAddrIdx As Long, ByRef r_wData As DSPWave, ByRef r_dData As Double) As Long
    On Error Resume Next
    Dim wTemp As New DSPWave
    Dim wIndex As New DSPWave

    wIndex.CreateRamp lAddrIdx * gD_slOTP_REGDATA_BW, 1, gD_slOTP_REGDATA_BW, DspLong
    wTemp = wDataBit.ConvertDataTypeTo(DspLong).Copy
    Call gD_wReadData.ReplaceElements(wIndex, wTemp)

    r_wData = wDataBit.ConvertStreamTo(tldspParallel, gD_slOTP_REGDATA_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspDouble)
    r_dData = r_wData.Element(0)
Exit Function
End Function

''''New 20191203 OTP-DSP
'___Return a decimal read value decided by the section of gD_wReadData
'Public Function otp_get_read_DataWave_LoopAddr(ByVal inWave As DSPWave, ByVal addrIdx_Start As Long, ByVal addrIdx_End As Long, ByRef rtnDataValWave As DSPWave) As Long ', ByRef r_dData As Double) As Long
'On Error Resume Next
'    Dim wTemp As New DSPWave
'    Dim wIndex As New DSPWave
'
'    wIndex.CreateRamp addrIdx_Start * gD_slOTP_REGDATA_BW, 1, (addrIdx_End - addrIdx_Start + 1) * gD_slOTP_REGDATA_BW, DspLong
'    wTemp = inWave.ConvertDataTypeTo(DspLong).Copy
'    Call gD_wReadData.ReplaceElements(wIndex, wTemp)
'
''    rtnDataValWave = inWave.ConvertStreamTo(tldspParallel, gD_slOTP_REGDATA_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspDouble)
''    r_dData = rtnDataValWave.Element(0)
'End Function
'20190513 This function will output the decimal value of the cut of defined BW in gD_wPGMData
Public Function otp_get_ConvStream_WriteData(ByVal lStartPnt As Long, ByVal lEndPnt As Long, ByRef r_dData As Double) As Double
    On Error Resume Next
    Dim wTemp As New DSPWave
    Dim wParaData As New DSPWave

    wTemp = gD_wPGMData.Select(lStartPnt, 1, lEndPnt - lStartPnt + 1).Copy
    wParaData = wTemp.ConvertStreamTo(tldspParallel, lEndPnt - lStartPnt + 1, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
    r_dData = wParaData.Element(0)
Exit Function
End Function
'20190513 This function will output the decimal value of the cut of defined BW in gD_wReadData
Public Function otp_get_ConvStream_ReadData(ByVal lStartPnt As Long, ByVal lEndPnt As Long, ByRef r_dData As Long) As Long
    On Error Resume Next
    Dim wTemp As New DSPWave
    Dim wParaData As New DSPWave

    wTemp = gD_wReadData.Select(lStartPnt, 1, lEndPnt - lStartPnt + 1).Copy
    wParaData = wTemp.ConvertStreamTo(tldspParallel, lEndPnt - lStartPnt + 1, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
    r_dData = wParaData.Element(0)
Exit Function
End Function
'20190513 This function will output the decimal value of the convert from gD_slOTP_REGDATA_BW in gD_wPGMData
Public Function otp_get_ConvStream_WriteRegData(ByVal lAddr As Long, ByRef r_dPGMData As Double) As Double
    On Error Resume Next
    Dim wTemp As New DSPWave
    Dim wParaData As New DSPWave

    wTemp = gD_wPGMData.Select(lAddr * gD_slOTP_REGDATA_BW, 1, gD_slOTP_REGDATA_BW).Copy
    wParaData = wTemp.ConvertStreamTo(tldspParallel, gD_slOTP_REGDATA_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspDouble)
    r_dPGMData = wParaData.Element(0)
Exit Function
End Function
'20190513 This function will output the decimal value convert from the cut of gD_slOTP_REGDATA_BW in gD_wReadData
Public Function otp_get_ConvStream_ReadRegData(ByVal lAddr As Long, ByRef r_dReadData As Double) As Double
    On Error Resume Next
    Dim wTemp As New DSPWave
    Dim wParaData As New DSPWave

    wTemp = gD_wReadData.Select(lAddr * gD_slOTP_REGDATA_BW, 1, gD_slOTP_REGDATA_BW).Copy
    wParaData = wTemp.ConvertStreamTo(tldspParallel, gD_slOTP_REGDATA_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspDouble)
    r_dReadData = wParaData.Element(0)
Exit Function
End Function

'#--------------------------------------------------
'# Calculate Look-Up-Table
'#--------------------------------------------------
Private Function otp_calculate_crc8_lut(ByVal lPolyInNormalNatation As Long) As Long
    On Error Resume Next
    Dim lByteIdx As Long
    Dim lCurrentByte As Long
    Dim lBitIdx As Long
    
    ''alCrc8Lut = []
    Dim alCrc8Lut(255) As Long
    Dim wCrc8Lut As New DSPWave
    wCrc8Lut.CreateConstant 0, 256, DspLong ''''because 8 bits => size=256

    ''for lByteIdx in range(0,256):
    For lByteIdx = 0 To 255
        '# Initialyze byte
        lCurrentByte = lByteIdx

        '# Compute CRC on this byte according to polynom
        ''for lBitIdx in range(0,8):
        For lBitIdx = 0 To 7
            ''if (lCurrentByte & 0x80) != 0:                # If MSB is set, shift and apply polynom
            If (lCurrentByte And &H80) <> 0 Then
                ''lCurrentByte = ((lCurrentByte*2) & 0xFF)
                lCurrentByte = ((lCurrentByte * 2) And &HFF)

                ''lCurrentByte = (lCurrentByte ^ lPolyInNormalNatation)
                lCurrentByte = (lCurrentByte Xor lPolyInNormalNatation)
            Else                                           '# If MSB is cleared, just shift
                ''lCurrentByte = ((lCurrentByte*2) & 0xFF)
                lCurrentByte = ((lCurrentByte * 2) And &HFF)
            End If
        Next lBitIdx

        '# Update LUT
        ''alCrc8Lut.append(lCurrentByte)
        alCrc8Lut(lByteIdx) = lCurrentByte
    Next lByteIdx

    wCrc8Lut.Data = alCrc8Lut
    gD_wCRCSelfLUT = wCrc8Lut.Copy
Exit Function
End Function

''''Put in OTP_initialize(), just do once
''''Call RunDSP.otp_Initialize_crc8(&HCF, &H0)    => MPXX
''''Call RunDSP.otp_Initialize_crc8(&H7, &H0)     => ML3T
'#--------------------------------------------------
'# Constructor
'# (polynom should be given in 'normal notation')
'#--------------------------------------------------
Public Function otp_Initialize_crc8(ByVal lPolyInNormalNatation As Long, ByVal lIniValue As Long) As Long
    On Error Resume Next
    ' Public variables (extracted from SPDS)
    ''self.value = lIniValue
    'self_value = lIniValue
     gD_slCRCSelfValue = lIniValue

    ''''gen Look Up Table DSPWave: gD_wCRCSelfLUT
    Call otp_calculate_crc8_lut(lPolyInNormalNatation)
Exit Function
End Function
'#--------------------------------------------------
'# Update CRC8 with a single byte of data
'#--------------------------------------------------
'
''def update_single_byte(self,byte):
Private Function XorPreviousCrcByte(ByVal lByte As Long) As Long
    On Error Resume Next
    Dim lData As Long

    ''lData = byte ^ self.value
    ''''lData = lByte Xor self_value
    lData = lByte Xor gD_slCRCSelfValue

    ''self.value = self.lut[lData]
    ''''self_value = self_lut(lData)
    gD_slCRCSelfValue = gD_wCRCSelfLUT.Element(lData)
Exit Function
End Function
'#--------------------------------------------------
'# Update CRC8 with a list of multiple bytes of data
'#--------------------------------------------------
'
''def update_bytes(self, byte_list):
''''Public Sub update_bytes(byte_list() As Long)
Public Function otp_CalculateCRC(ByVal wByteList As DSPWave, ByRef r_CrcResult As Long) As Long
    On Error Resume Next
    Dim lIdx As Long
    Dim lByte As Long
    Dim alList() As Long
    
    alList = wByteList.Data
    For lIdx = 0 To UBound(alList)
        lByte = alList(lIdx)
        Call XorPreviousCrcByte(lByte)
    Next lIdx
    
    ''''return the result of CRC8
    r_CrcResult = gD_slCRCSelfValue
    gD_slCRCSelfValue = 0 'Clean result
Exit Function
End Function

''''20190618 update
'___Replace elements in gD_wPGMData by the input value.
Public Function LocateOTPData2gDw(ByVal lInputValue As Long, ByVal lBitWidth As Long, ByVal wIndex As DSPWave) As Long
    On Error Resume Next
    Dim wBinary As New DSPWave
    Dim wInputValue As New DSPWave
  
    wInputValue.CreateConstant lInputValue, 1, DspLong
    wBinary = wInputValue.ConvertStreamTo(tldspSerial, lBitWidth, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
    gD_wPGMData.ReplaceElements wIndex, wBinary
Exit Function
End Function

Public Function SumUpCatReadData(ByRef r_lSumReadData As Long) As Long
    On Error Resume Next
    Dim rtnBlankChkDataWave As New DSPWave

    r_lSumReadData = gD_wReadData.CalcSum

Exit Function
End Function
'JY 20200406 Fuji's new request
Public Function SumUpCatReadData_ExceptECID(ByVal StarAddr As Long, ByVal SumCnt As Long, ByRef r_lSumReadData As Long) As Long
    On Error Resume Next
    Dim rtnBlankChkDataWave As New DSPWave
    rtnBlankChkDataWave = gD_wReadData.Select(StarAddr, 1, SumCnt)
    r_lSumReadData = rtnBlankChkDataWave.CalcSum
Exit Function
End Function

Public Function otp_compare_PGM_Read_DataWave(ByRef r_wPGMDataByBW As DSPWave, ByRef r_wReadDataByBW As DSPWave, ByRef r_wCompare As DSPWave) As Long
    On Error Resume Next
    r_wCompare.CreateConstant -1, gD_wPGMData.SampleSize, DspLong

    ''''because gD_slOTP_REGDATA_BW is constant, i.e., =32 bits.
    r_wPGMDataByBW = gD_wPGMData.ConvertStreamTo(tldspParallel, gD_slOTP_REGDATA_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspDouble)
    r_wReadDataByBW = gD_wReadData.ConvertStreamTo(tldspParallel, gD_slOTP_REGDATA_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspDouble)
    r_wCompare = r_wReadDataByBW.LogicalCompare(EqualTo, r_wPGMDataByBW).ConvertDataTypeTo(DspLong) 'Equal to is "1", non-equal to is "0"
    
    r_wCompare = r_wCompare.Multiply(-1) 'Convert it to true (-1) and false(0)
    ''''if PGM=Read, rtnCmpSum=0
    'rtnCmpSum = r_wCompare.CalcSum
    
Exit Function
End Function

Public Function otp_Read_DataWave(ByRef r_wReadDataByBW As DSPWave) As Long
    On Error Resume Next

    r_wReadDataByBW = gD_wReadData.ConvertStreamTo(tldspParallel, gD_slOTP_REGDATA_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspDouble)
    
Exit Function
End Function
'20191106 Divide by 8 bits
Public Function otp_allotECID(ByRef r_wEcidPgmDataByBW As DSPWave) As Long
    On Error Resume Next

    r_wEcidPgmDataByBW = gD_wDEIDPGMBits.ConvertStreamTo(tldspParallel, 8, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
    
Exit Function
End Function

''''20191008, update method with addr loop inside
''''update method for the FTProg to A0 except Lockbit to prevent burn '1' again
'___Concatenate the addr and data dspwave
Public Function otp_get_pgm_AddrDataWave_maskECID_LoopAddr(ByVal sbOtpedECID As Boolean, ByVal wPgmWaveMaskECID As DSPWave, ByVal lOtpOfs As Long, _
                                                           ByVal laddrStart As Long, ByVal laddrEnd As Long, _
                                                           ByRef r_wAllAddrAndDataBit As DSPWave) As Long
''''20200313 MP7P, use ECID_Mask() outside
'''Public Function otp_get_pgm_AddrDataWave_maskECID_LoopAddr(ByVal lOtpOfs As Long, _
'''                                                           ByVal laddrStart As Long, ByVal laddrEnd As Long, _
'''                                                           ByRef r_wAllAddrAndDataBit As DSPWave) As Long
    On Error Resume Next
    Dim lAddrIdx As Long
    Dim lAddrwiOfs As Long
    Dim wAddr As New DSPWave
    Dim wAddrBit As New DSPWave
    Dim wDataBit As New DSPWave
    Dim lSelectedStartBit As Long
    Dim wAddrAndData As New DSPWave
    Dim wIndex As New DSPWave
    Dim wTemp As New DSPWave
    
    Dim lTotalAddrs As Long ''''numbers of all address from 0 to laddrEnd
    Dim wAddrDataBW As Long ''''numbers of bits (addr+data)

    wAddrDataBW = gD_slOTP_ADDR_BW + gD_slOTP_REGDATA_BW
    lTotalAddrs = (laddrEnd - 0) + 1 ''''(laddrEnd - laddrStart) + 1
    
    ''''[NOTICE] Here MUST gen all bits from addr=0 to laddrEnd,
    ''''then it will be easy to be replaced later on (just be convient mathematically)
    r_wAllAddrAndDataBit.CreateConstant 0, wAddrDataBW * lTotalAddrs, DspLong

    wTemp = gD_wPGMData.Copy
    
''''20200313 MP7P use ECID_Mask() outside, need to check
    '''' sbOtpedECID = True  , means ECID burned     => g_sbOtpedECID = True
    '''' sbOtpedECID = False , means ECID not burned => g_sbOtpedECID = False
    If sbOtpedECID Then
        ''''because ECID has burned, so we set all ECID bits as "0"
        wTemp = wPgmWaveMaskECID.Copy
    Else
        wTemp = gD_wPGMData.Copy
    End If

    For lAddrIdx = laddrStart To laddrEnd
        lAddrwiOfs = lAddrIdx + lOtpOfs
        wAddr.CreateConstant lAddrwiOfs, 1, DspLong
        wAddrBit = wAddr.ConvertStreamTo(tldspSerial, gD_slOTP_ADDR_BW, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
        
        lSelectedStartBit = lAddrIdx * gD_slOTP_REGDATA_BW
        wAddrAndData.CreateConstant 0, gD_slOTP_REGDATA_BW, DspLong

        wDataBit = wTemp.Select(lAddrIdx * gD_slOTP_REGDATA_BW, 1, gD_slOTP_REGDATA_BW).Copy
        wAddrAndData = wAddrBit.Concatenate(wDataBit)
        
        ''''create the index to replace the elements in the specific location
        wIndex.CreateRamp lAddrIdx * wAddrDataBW, 1, wAddrDataBW, DspLong
        
        r_wAllAddrAndDataBit.ReplaceElements wIndex, wAddrAndData
    Next lAddrIdx

Exit Function
End Function

