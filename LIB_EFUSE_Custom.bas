Attribute VB_Name = "LIB_EFUSE_Custom"

Option Explicit

Public gS_pre_eFuse_sheetName As String    ''''20160818 add
Public gS_pre_cfgtable_sheetName As String ''''20160901 add
Public Const gS_eFuse_sheetName = "EFUSE_BitDef_Table"  ''''"EFUSE_BitDef_Table","EFUSE_BitDef_Table_2bit", "EFUSE_BitDef_Table_2bit_wCond"
Public Const gS_cfgTable_sheetName = "CFG_Condition_Table"         ''''"Config_Table" ''''20170630
Public Const gS_cfgTable_SVM_sheetName = "CFG_Condition_Table_SVM" ''''"Config_Table_SVM"
Public Const gC_CFGSVM_BIT = 57   ''''64bits cfgtable only
Public Const gS_BKM_Name = "BKM_Info" '''20191221 BKM
Public gS_BKM_Unknown As String

''20191230 , make bank size to be dynamic
Public Const EcidCharPerLotId = 6
Public Const EcidBitPerLotIdChar = 6
Public Const EConfigReadBitWidth = 32
Public Const EcidReadBitWidth = 32
Public Const MONITORReadBitWidth = 32


Public gL_Sim_FuseBits() As Long      ''''it's used for the simulation.
Public gL_ECID_Sim_FuseBits() As Long ''''it's used for the simulation.
Public gL_CFG_Sim_FuseBits() As Long  ''''it's used for the simulation.

''''In the PinMap sheet, its order is Q31,Q30,Q29,......,Q2 ,Q1 ,Q0  ==> MSB
''''In the PinMap sheet, its order is Q0 ,Q1 ,Q2 ,......,Q29,Q30,Q31 ==> LSB
Public Const gC_eFuse_DigCap_BitOrder = "MSB"

''''<Important>-------------------------------------------------
''''The below Constant is decided from USI/USO DSSC pattern
Public Const gC_USI_DSSCRepeatCyclePerBit = 1
''Public Const gC_USO_DSSCRepeatCyclePerBit = 1 ''unused
Public gL_USI_DigSrcBits_Num As Long
Public gL_USO_DigCapBits_Num As Long
Public gL_USI_PatBitOrder As String  ''''CMN[bitLast...bit0] =>MSB...LSB
Public gL_USO_PatBitOrder As String  ''''CMN[bitLast...bit0] =>MSB...LSB
Public gL_CMP_DigCapBits_Num As Long
Public gS_CMP_PatBitOrder As String  ''''CMN[bitLast...bit0] =>MSB...LSB or CMN[bit0...bitLast] =>LSB...MSB
''''------------------------------------------------------------
''''20171103 add
Public Const gC_UDRE_USI_DSSCRepeatCyclePerBit = 1
Public gL_UDRE_USI_DigSrcBits_Num As Long
Public gL_UDRE_USO_DigCapBits_Num As Long
Public gL_UDRE_USI_PatBitOrder As String  ''''CMN[bitLast...bit0] =>MSB...LSB
Public gL_UDRE_USO_PatBitOrder As String  ''''CMN[bitLast...bit0] =>MSB...LSB
Public gL_CMPE_DigCapBits_Num As Long
Public gS_CMPE_PatBitOrder As String      ''''CMN[bitLast...bit0] =>MSB...LSB or CMN[bit0...bitLast] =>LSB...MSB
''''------------------------------------------------------------
Public Const gC_UDRP_USI_DSSCRepeatCyclePerBit = 1
Public gL_UDRP_USI_DigSrcBits_Num As Long
Public gL_UDRP_USO_DigCapBits_Num As Long
Public gL_UDRP_USI_PatBitOrder As String  ''''CMN[bitLast...bit0] =>MSB...LSB
Public gL_UDRP_USO_PatBitOrder As String  ''''CMN[bitLast...bit0] =>MSB...LSB
Public gL_CMPP_DigCapBits_Num As Long
Public gS_CMPP_PatBitOrder As String      ''''CMN[bitLast...bit0] =>MSB...LSB or CMN[bit0...bitLast] =>LSB...MSB
''''------------------------------------------------------------
Public gL_SEN_CRC_EndBit As Long
Public gS_SEN_CRC_Stage As String
Public gL_SEN_CRC_LSBbit As Long        '''' 20161026 ADD CRC
Public gL_SEN_CRC_MSBbit As Long        '''' 20161026 ADD CRC
Public gL_SEN_CRC_BitWidth As Long
Public gS_MON_CRC_Stage As String
Public gL_MON_CRC_EndBit As Long
Public gL_MON_CRC_LSBbit As Long        '''' 20161026 ADD CRC
Public gL_MON_CRC_MSBbit As Long        '''' 20161026 ADD CRC
Public gL_MON_CRC_BitWidth As Long
Public gS_UID_CRC_Stage As String
Public gL_UID_CRC_EndBit As Long
Public gL_UID_CRC_LSBbit As Long        '''' 20161026 ADD CRC
Public gL_UID_CRC_MSBbit As Long        '''' 20161026 ADD CRC
Public gL_UID_CRC_BitWidth As Long

Public gDW_SEN_CRC_calcBits_Temp As New DSPWave
Public gDW_MON_CRC_calcBits_Temp As New DSPWave
Public gDW_UID_CRC_calcBits_Temp As New DSPWave

Public gL_ECID_CRC_EndBit As Long       '''' Was 64bits only, Const gL_ECID_CRC_EndBit = 63
Public gS_ECID_CRC_Stage As String      '''' 20161003 ADD CRC
Public gL_ECID_CRC_LSB As Long          '''' 20161003 ADD CRC
Public gL_ECID_CRC_MSB As Long          '''' 20161003 ADD CRC
Public gL_ECID_CRC_BitWidth As Long
Public gS_CFG_CRC_Stage As String       '''' 20161003 ADD CRC
Public gL_CFG_CRC_LSBbit As Long        '''' 20161026 ADD CRC
Public gL_CFG_CRC_MSBbit As Long        '''' 20161026 ADD CRC
Public gL_CFG_CRC_BitWidth As Long
Public gS_ECID_CRC_PgmFlow As String
Public gL_ECID_CRC_calcBits() As Long   '''' 20170823 ECID CRC calcBits Array, =1 means CRC calculated bit, =0 means CRC ignore bit
Public gL_CFG_CRC_calcBits() As Long    '''' 20170823  CFG CRC calcBits Array, =1 means CRC calculated bit, =0 means CRC ignore bit

Public gS_ECID_Read_calcCRC_hexStr As New SiteVariant
Public gS_ECID_Read_calcCRC_bitStrM As New SiteVariant

''''201812XX add
Public gS_CFG_Read_calcCRC_hexStr As New SiteVariant
Public gS_CFG_Read_calcCRC_bitStrM As New SiteVariant

Public gS_SEN_Read_calcCRC_hexStr As New SiteVariant
Public gS_SEN_Read_calcCRC_bitStrM As New SiteVariant

Public gS_MON_Read_calcCRC_hexStr As New SiteVariant
Public gS_MON_Read_calcCRC_bitStrM As New SiteVariant

Public gS_UID_Read_calcCRC_hexStr As New SiteVariant
Public gS_UID_Read_calcCRC_bitStrM As New SiteVariant


Public Function auto_ECIDConstant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ECIDConstant_Initialize"

    Dim i As Long, j As Long
    Dim k As Long, n As Long
    Dim idx As Long
    Dim ss As Variant
    Dim efflastBit As Long
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_defval As Variant
    Dim m_len As Long
    Dim m_defreal As String
    Dim m_bitwidth As Long
    Dim m_defvalhex As String ''''without prefix "0x"

    Dim m_cmt As String ''''20170823
    Dim m_cmtArr() As String
    Dim m_tmpArr() As String
    Dim m_bit_min As Long
    Dim m_bit_max As Long
    Dim m_value As Long
    Dim m_find_ECID_calcBits_flag As Boolean
    Dim m_binarr() As Long
    Dim ms_defval As New SiteVariant

    ''''Get max length of category name
    gI_ECID_catename_maxLen = 0
    For i = 0 To UBound(ECIDFuse.Category)
        m_len = Len(ECIDFuse.Category(i).Name)
        
        If (m_len > gI_ECID_catename_maxLen) Then
            gI_ECID_catename_maxLen = m_len
        End If
    Next i
    gI_ECID_catename_maxLen = gI_ECID_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------
    
    ''20191230 , make bank size to be dynamic
    Dim m_EffectiveBit As Long
    If (UCase(ECIDFuse.Category(0).MSBFirst) = "Y") Then
        m_EffectiveBit = ECIDFuse.Category(UBound(ECIDFuse.Category) - 1).LSBbit + 1
    Else
        m_EffectiveBit = ECIDFuse.Category(UBound(ECIDFuse.Category) - 1).MSBbit + 1
    End If
    
    EcidWriteBitExpandWidth = 1
    'EcidReadBitWidth = 32
    
    If (gS_EFuse_Orientation = "UP2DOWN") Then
        EcidBlock = 2
        EcidBitsPerRow = EcidReadBitWidth
        EcidRowPerBlock = m_EffectiveBit / EcidBitsPerRow

        EcidReadCycle = EcidRowPerBlock * EcidBlock            ''=8*2=16 <Notice>

    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then ''''2-Bit
        EcidBlock = 2
        EcidBitsPerRow = EcidReadBitWidth / EcidBlock
        EcidRowPerBlock = m_EffectiveBit / EcidBitsPerRow

        EcidReadCycle = EcidRowPerBlock


    ElseIf (gS_EFuse_Orientation = "SingleUp") Then ''''1-Bit
        EcidBlock = 1
        EcidBitsPerRow = EcidReadBitWidth
        EcidRowPerBlock = m_EffectiveBit / EcidBitsPerRow
        
        EcidReadCycle = EcidRowPerBlock * EcidBlock            ''=16*1=16 <Notice>                  ''=32


    ''''Below is reserved for the future, there is NO any definition at present.
    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
    End If
    
    EcidBitPerBlockUsed = EcidRowPerBlock * EcidBitsPerRow ''=16x16=256
    EcidBitPerBlock = EcidBitPerBlockUsed
    ECIDTotalBits = EcidBitPerBlockUsed * EcidBlock        ''=256x2=512
    ECIDBitPerCycle = EcidReadBitWidth
    
'    If (gS_EFuse_Orientation = "UP2DOWN") Then
'        EcidBlock = 2
'        EcidRowPerBlock = 8
'        EcidBitsPerRow = 32
'        EcidWriteBitExpandWidth = 120
'        EcidReadBitWidth = 32
'
'        EcidBitPerBlockUsed = EcidRowPerBlock * EcidBitsPerRow ''=8x32=256
'        EcidBitPerBlock = EcidBitPerBlockUsed                  ''=256
'        EcidReadCycle = EcidRowPerBlock * EcidBlock            ''=8*2=16 <Notice>
'        ECIDTotalBits = EcidBitPerBlockUsed * EcidBlock        ''=256x2=512
'        ECIDBitPerCycle = EcidReadBitWidth                     ''=32
'        EcidHiLimitSingleDoubleBitCheck = 0
'
'    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then ''''2-Bit
'        EcidBlock = 2
'        EcidRowPerBlock = 16
'        EcidBitsPerRow = 16
'        EcidWriteBitExpandWidth = 120
'        EcidReadBitWidth = 32
'
'        EcidBitPerBlockUsed = EcidRowPerBlock * EcidBitsPerRow ''=16x16=256
'        EcidBitPerBlock = EcidBitPerBlockUsed                  ''=256
'        EcidReadCycle = EcidRowPerBlock                        ''=16
'        ECIDTotalBits = EcidBitPerBlockUsed * EcidBlock        ''=256x2=512
'        ECIDBitPerCycle = EcidReadBitWidth                     ''=32
'        EcidHiLimitSingleDoubleBitCheck = 0
'
'    ElseIf (gS_EFuse_Orientation = "SingleUp") Then ''''1-Bit
'        EcidBlock = 1
'        EcidRowPerBlock = 16
'        EcidBitsPerRow = 32
'        EcidWriteBitExpandWidth = 120 ''''<based on pattern>
'        EcidReadBitWidth = 32
'
'        EcidBitPerBlockUsed = EcidRowPerBlock * EcidBitsPerRow ''=16x32=512
'        EcidBitPerBlock = EcidBitPerBlockUsed                  ''=512
'        EcidReadCycle = EcidRowPerBlock * EcidBlock            ''=16*1=16 <Notice>
'        ECIDTotalBits = EcidBitPerBlockUsed * EcidBlock        ''=512x1=512
'        ECIDBitPerCycle = EcidReadBitWidth                     ''=32
'        EcidHiLimitSingleDoubleBitCheck = 0
'
'    ''''Below is reserved for the future, there is NO any definition at present.
'    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
'    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
'    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
'    End If

    
    idx = ECIDIndex("Lot_ID")
    If (idx >= 0) Then
        LOTID_FIRST_BIT = ECIDFuse.Category(idx).MSBbit
        LOTID_LAST_BIT = ECIDFuse.Category(idx).LSBbit
        LOTID_BITWIDTH = ECIDFuse.Category(idx).BitWidth
    End If
    
    idx = ECIDIndex("Wafer_ID")
    If (idx >= 0) Then
        WAFERID_FIRST_BIT = ECIDFuse.Category(idx).MSBbit
        WAFERID_LAST_BIT = ECIDFuse.Category(idx).LSBbit
        WAFERID_BITWIDTH = ECIDFuse.Category(idx).BitWidth
    End If

    idx = ECIDIndex("X_Coordinate")
    If (idx >= 0) Then
        XCOORD_FIRST_BIT = ECIDFuse.Category(idx).MSBbit
        XCOORD_LAST_BIT = ECIDFuse.Category(idx).LSBbit
        XCOORD_BITWIDTH = ECIDFuse.Category(idx).BitWidth
        If (UCase(ECIDFuse.Category(idx).LoLMT) = "N/A") Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: X_Coordinate, please have a correct Low Limit value."
            ECIDFuse.Category(idx).LoLMT = 0
        End If
        If (UCase(ECIDFuse.Category(idx).HiLMT) = "N/A") Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: X_Coordinate, please have a correct High Limit value."
            ECIDFuse.Category(idx).HiLMT = 0
        End If
        XCOORD_LoLMT = CLng(ECIDFuse.Category(idx).LoLMT)
        XCOORD_HiLMT = CLng(ECIDFuse.Category(idx).HiLMT)
    End If

    idx = ECIDIndex("Y_Coordinate")
    If (idx >= 0) Then
        YCOORD_FIRST_BIT = ECIDFuse.Category(idx).MSBbit
        YCOORD_LAST_BIT = ECIDFuse.Category(idx).LSBbit
        YCOORD_BITWIDTH = ECIDFuse.Category(idx).BitWidth
        If (UCase(ECIDFuse.Category(idx).LoLMT) = "N/A") Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: Y_Coordinate, please have a correct Low Limit value."
            ECIDFuse.Category(idx).LoLMT = 0
        End If
        If (UCase(ECIDFuse.Category(idx).HiLMT) = "N/A") Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: Y_Coordinate, please have a correct High Limit value."
            ECIDFuse.Category(idx).HiLMT = 0
        End If
        YCOORD_LoLMT = CLng(ECIDFuse.Category(idx).LoLMT)
        YCOORD_HiLMT = CLng(ECIDFuse.Category(idx).HiLMT)
    End If
    
    gI_Index_DEID = ECIDIndex("ECID_DEID")
    gS_ECID_CRC_Stage = "" ''''initial
    
    ''''20160103 Add for TestChip (only ECID)
    For i = 0 To UBound(ECIDFuse.Category)
        m_catename = ECIDFuse.Category(i).Name
        m_algorithm = LCase(ECIDFuse.Category(i).algorithm)
        m_defval = ECIDFuse.Category(i).DefaultValue
        m_defreal = LCase(ECIDFuse.Category(i).Default_Real)
        m_bitwidth = ECIDFuse.Category(i).BitWidth

        If (m_algorithm = "crc") Then m_defval = 0
        ''''20171211, Here it's used to judge default value if is Hex or Binary
        m_defval = auto_checkDefaultValue(m_defval, m_algorithm, m_binarr, m_bitwidth, m_defreal)
        'm_defval = auto_checkDefaultValue(m_defval, m_binarr, m_bitwidth, m_defreal)
        ECIDFuse.Category(i).DefValBitArr = m_binarr
        ms_defval = m_defval
        
        If (m_algorithm = "fuse" Or m_algorithm = "device") Then
            ''''Actually, it should be "Default" but sometimes Fuji put it as "Real" in the Algorithm column
'            If (m_algorithm = "real") Then
'                For Each ss In TheExec.Sites
'                    Call auto_eFuse_SetWriteDecimal("ECID", m_catename, m_defval, False)
'                Next ss
'            End If

        '''' 20161003 ADD CRC <Be Carefully here, Need to Check>
        ElseIf (m_algorithm = "crc") Then
            ''''<NOTICE>20170331, ECID CRC always coding/fusing as [MSB......LSB]
            gS_ECID_CRC_Stage = LCase(ECIDFuse.Category(i).Stage)
            gL_ECID_CRC_BitWidth = ECIDFuse.Category(i).BitWidth
            If (ECIDFuse.Category(i).MSBbit < ECIDFuse.Category(i).LSBbit) Then
                ''''current case in eFuse ECID table
                gL_ECID_CRC_LSB = ECIDFuse.Category(i).MSBbit ''''Tmys
                gL_ECID_CRC_MSB = ECIDFuse.Category(i).LSBbit ''''Tmys
            Else
                ''''MSB bit location > LSB bit location
                gL_ECID_CRC_LSB = ECIDFuse.Category(i).LSBbit
                gL_ECID_CRC_MSB = ECIDFuse.Category(i).MSBbit
            End If
            
            ''''20170815 update
            ''''20170915 update to support NOT continuous bits and have some prevent methods.
            ''''----------------------------------------------------------------------------------------
            ''''Example
            ''''----------------------------------------------------------------------------------------
            ''''Ex: Comment/Description:: PgmFlow=nonDEID, CRC_calcBits=[255:0]
            ''''----------------------------------------------------------------------------------------
            m_cmt = Trim(UCase(ECIDFuse.Category(i).comment))
            m_cmtArr = Split(m_cmt, ",")
            gL_ECID_CRC_EndBit = 63       ''''default
            gS_ECID_CRC_PgmFlow = "DEID"  ''''default
            ReDim gL_ECID_CRC_calcBits(EcidBitPerBlockUsed - 1) ''''<MUST>be here

            m_find_ECID_calcBits_flag = False ''''default
            
            For j = 0 To UBound(m_cmtArr)
                ''''reuse the variable
                m_cmt = Trim(UCase(m_cmtArr(j)))
                m_cmt = Replace(m_cmt, "[", "", 1)
                m_cmt = Replace(m_cmt, "]", "", 1)
                m_bit_min = 99999
                m_bit_max = -99999

                If (m_cmt Like UCase("*calcBits*=*[*:*]*")) Then ''''<MUST>
                    m_find_ECID_calcBits_flag = True
                End If
                
                If (m_cmt Like UCase("*PgmFlow*")) Then
                    m_cmt = Trim(Replace(m_cmt, "PgmFlow", "", 1, 1, vbTextCompare))
                    m_cmt = Trim(Replace(m_cmt, "=", "", 1, 1, vbTextCompare))
                    gS_ECID_CRC_PgmFlow = m_cmt

                ElseIf (m_cmt Like UCase("*:*")) Then
                    ''''reuse the variable
                    m_cmt = Replace(m_cmt, "[", "", 1)
                    m_cmt = Replace(m_cmt, "]", "", 1)
                    m_bit_min = 99999
                    m_bit_max = -99999
                    If (m_cmt Like UCase("*=*")) Then
                        m_tmpArr = Split(m_cmt, "=")
                        m_cmt = Trim(m_tmpArr(1))
                    End If
                    m_tmpArr = Split(m_cmt, ":")
                    For k = 0 To UBound(m_tmpArr)
                        m_cmt = Trim(m_tmpArr(k))
                        If (IsNumeric(m_cmt)) Then
                            m_value = CLng(m_cmt)
                            If (m_value > m_bit_max) Then
                                m_bit_max = m_value
                            Else
                                m_bit_min = m_value
                            End If
                        Else
                            TheExec.AddOutput funcName + ":: bits number is not a Numeric (" + m_cmt + ")"
                            GoTo errHandler
                        End If
                    Next k
                    For n = m_bit_min To m_bit_max
                        gL_ECID_CRC_calcBits(n) = 1
                    Next n
                    gL_ECID_CRC_EndBit = m_bit_max ''''<MUST>
                Else
                    ''''Unexpected Format Case
                    TheExec.AddOutput funcName + ":: Unexpected Formaton ECID CRC Comment element: " + m_cmtArr(j)
                    TheExec.Datalog.WriteComment funcName + ":: Unexpected Formaton ECID CRC Comment element: " + m_cmtArr(j)
                    ''GoTo errHandler
                End If
            Next j
            
            If (m_find_ECID_calcBits_flag = False) Then
                ''''Exception Case
                TheExec.AddOutput funcName + ":: Problem on ECID CRC Comment Format, no calcBits definition, " + ECIDFuse.Category(i).comment
                TheExec.Datalog.WriteComment funcName + ":: Problem on ECID CRC Comment Format, no calcBits definition, " + ECIDFuse.Category(i).comment
                GoTo errHandler ''''<MUST>
            End If
        End If

        ''''20160328 update, 20160624 update
        If (m_defreal <> "real" And m_defreal <> "bincut") Then
            ECIDFuse.Category(i).DefaultValue = m_defval
            ECIDFuse.Category(i).HiLMT = m_defval
            ECIDFuse.Category(i).LoLMT = m_defval

            ''''20180712, will do the update later on for "vddbin: safe voltage"
'            If (m_algorithm = "vddbin" And m_defreal Like "safe*voltage") Then
'            Else
'                For Each ss In TheExec.Sites.Existing
'                    Call auto_eFuse_SetWriteDecimal("ECID", m_catename, m_defval, False)
'                Next ss
'            End If
'        Else
'            ''''20180723, For "Real", give an initial value as m_defval, usually it is "Zero".
'            For Each ss In TheExec.Sites.Existing
'                Call auto_eFuse_SetWriteDecimal("ECID", m_catename, m_defval, False)
'            Next ss
        End If
        
        If (m_algorithm = "vddbin" And m_defreal Like "safe*voltage") Then
        Else
             Call auto_eFuse_SetWriteVariable_SiteAware("ECID", m_catename, ms_defval, False)
        End If
    Next i

    ''''initialize the below variable for each run (initFlows)
    For Each ss In TheExec.sites.Existing
        HramLotId(ss) = ""
        HramWaferId(ss) = 0
        HramXCoord(ss) = -32768
        HramYCoord(ss) = -32768
        gS_ECID_CRC_HexStr(ss) = "" ''''MUST be, 20161004 update
    Next ss

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_CFGConstant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CFGConstant_Initialize"

    Dim i As Long
    Dim j As Long
    Dim k As Long, n As Long, kk As Long
    Dim ss As Variant
    Dim m_stage As String
    Dim m_pkgname As String
    Dim m_algorithm As String
    Dim m_catename As String
    Dim m_defval As Variant
    Dim m_MSBBit As Long
    Dim m_len As Long
    Dim m_resolution As Double
    Dim m_defreal As String
    Dim m_bitwidth As Long
    Dim m_defvalhex As String ''''without prefix "0x"

    Dim m_cmt As String ''''20170823
    Dim m_cmtArr() As String
    Dim m_tmpArr() As String
    Dim m_bit_min As Long
    Dim m_bit_max As Long
    Dim m_value As Long
    Dim m_find_CFG_calcBits_flag As Boolean
    Dim m_binarr() As Long

    Dim m_decimal As Long
    Dim Count As Long
    Dim m_condIdx As Long
    Dim m_bitStrM As String
    Dim m_cmpValue As Long
    Dim m_cmpValue_sum As Long
    Dim m_condCateCNT As Long
    Dim ms_defval As New SiteVariant
    
    
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_lolmt_cateArr() As Variant
    Dim m_hilmt_cateArr() As Variant
    ReDim m_lolmt_cateArr(UBound(CFGFuse.Category))
    ReDim m_hilmt_cateArr(UBound(CFGFuse.Category))
    
    Dim m_idx0 As Long
    Dim m_idx1 As Long
    Dim m_divider As Long
    Dim m_LSBbit As Long

    ''''Get max length of category name
    gI_CFG_catename_maxLen = 0
    For i = 0 To UBound(CFGFuse.Category)
        m_len = Len(CFGFuse.Category(i).Name)
        
        If (m_len > gI_CFG_catename_maxLen) Then
            gI_CFG_catename_maxLen = m_len
        End If
    Next i
    gI_CFG_catename_maxLen = gI_CFG_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------

    ''''initialize
    ''''FT3 could be different, check in auto_GetSiteFlagName()
    ''''CP=A00, check in auto_Copy_CFGTable_Data_to_Array_byStage()
    ''''201811XX Update [leave to User's maintain]
    ''''         "CFG_1ST_SI" means 1st Silicon as the test plan
    ''''-----------------------------------------------------------------------------------
    gS_cfgFlagname = "ALL_0" ''or "1ST_SI"  ''''default all zero value, 20160103 update
    ''''-----------------------------------------------------------------------------------
    Count = 0 'Initialization
    Call auto_GetSiteFlagName(Count, gS_cfgFlagname, True)
    gS_cfgFlagname = UCase(gS_cfgFlagname)
    If (Count = 0) Then
        gS_cfgFlagname = "ALL_0" ''''could be "1ST_SI" or "A00"
    ElseIf (Count <> 1) Then
        TheExec.Datalog.WriteComment vbCrLf & "<WARNING> There are more one CFG condition Flag selected. Please check it!! " ''''20160927 add
        gS_cfgFlagname = "ALL_0"
        GoTo errHandler
        ''TheExec.Flow.TestLimit resultVal:=Count, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Error" ''''set fail
    End If
    kk = CFGTabIndex(gS_cfgFlagname)
    If (kk <> -1) Then
        gS_CFGCondTable_bitsStr = CFGTable.Category(kk).BitStrM ''was .BitStrM_byStage
    Else
        ''''use "ALL_0" as the presentive
        gS_CFGCondTable_bitsStr = CFGTable.Category(0).BitStrM ''was .BitStrM_byStage
    End If
    gS_cfgFlagname_pre = gS_cfgFlagname
    ''''-----------------------------------------------------------------------------------


    ''''-----------------------------------------------------------------------------------
    ''''20160728 update for the special tracker CFG_SVM A00 on CP1 only, CP2 can be tested.
    ''''<NOTICE> User needs to take case of the risk here
    ''''<Important> set True is for the special tracker version
''    If (gB_CFGSVM_A00_CP1 = True) Then
''        If (gB_CFG_SVM = True And gS_JobName Like "cp*") Then
''            TheExec.Flow.EnableWord("CFG_A00") = True
''            For Each ss In TheExec.Sites.Existing
''                TheExec.Sites.Item(ss).FlagState("A00") = logicTrue ''''<MUST>
''            Next ss
''            gS_cfgFlagname = "A00" ''''<MUST> be here
''            CFGFuse.Category(CFGIndex("CFG_Condition")).Stage = "CP1" ''''<MUST>
''            TheExec.Datalog.WriteComment funcName + ":: EnableWord CFG_SVM = True"
''            TheExec.Datalog.WriteComment funcName + ":: EnableWord CFG_A00 = True, gS_cfgFlagname = A00"
''            TheExec.Datalog.WriteComment funcName + ":: gB_CFGSVM_A00_CP1 = True"
''        End If
''    End If
    ''''-----------------------------------------------------------------------------------

    ''''When using Array(), must declare as Variant
    gS_DevRevArr = Array("A0", "A1", "A2", "A3", "A4", "A5", "", "", "B0", "B1", "B2", "B3", "B4", "B5", "", "", "C0", "C1") 'from test plan
    gS_Major_DevRevArr = Array("A", "B", "C", "D", "E", "F", "G", "H")
    
    ''20191230 , make bank size to be dynamic
    Dim m_EffectiveBit As Long
    m_EffectiveBit = CFGFuse.Category(UBound(CFGFuse.Category)).MSBbit + 1
    EConfig_Repeat_Cyc_for_Pgm = 1
    'EConfigReadBitWidth = 32
    If (gS_EFuse_Orientation = "UP2DOWN") Then
    
        EConfigBlock = 2
        EConfigBitsPerRow = EConfigReadBitWidth
        EConfigRowPerBlock = m_EffectiveBit / EConfigBitsPerRow '16
        EConfigReadCycle = EConfigRowPerBlock * EConfigBlock
         
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then

        EConfigBlock = 2
        EConfigBitsPerRow = EConfigReadBitWidth / EConfigBlock   ''''here must be 16
        EConfigRowPerBlock = m_EffectiveBit / EConfigBitsPerRow '16
        EConfigReadCycle = EConfigRowPerBlock
        
    ElseIf (gS_EFuse_Orientation = "SingleUp") Then

        EConfigBlock = 1
        EConfigBitsPerRow = EConfigReadBitWidth '32
        EConfigRowPerBlock = m_EffectiveBit / EConfigBitsPerRow '16
        EConfigReadCycle = EConfigRowPerBlock

    ''''it's reserved.
    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
    End If

    
    EConfigBitPerBlockUsed = EConfigRowPerBlock * EConfigBitsPerRow ''=16*32=512
    EConfigTotalBitCount = EConfigBlock * EConfigBitPerBlockUsed    ''=2*512=1024


'    If (gS_EFuse_Orientation = "UP2DOWN") Then
'
'        EConfigBlock = 2
'        EConfigBitsPerRow = 32
'        EConfigReadBitWidth = 32
'        EConfigRowPerBlock = 16
'        EConfigReadCycle = EConfigRowPerBlock * EConfigBlock
'
'        EConfigBitPerBlockUsed = EConfigRowPerBlock * EConfigBitsPerRow ''=16*32=512
'        EConfigTotalBitCount = EConfigBlock * EConfigBitPerBlockUsed    ''=2*512=1024
'
'        EConfig_Repeat_Cyc_for_Pgm = 1    '120
'        EConfigHiLimitSingleDoubleBitCheck = 0
'
'    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
'
'        EConfigBlock = 2
'        EConfigBitsPerRow = 16   ''''here must be 16
'        EConfigReadBitWidth = 32
'        EConfigRowPerBlock = 32 '16
'        EConfigReadCycle = EConfigRowPerBlock
'
'        EConfigBitPerBlockUsed = EConfigRowPerBlock * EConfigBitsPerRow ''=16*16=256
'        EConfigTotalBitCount = EConfigBlock * EConfigBitPerBlockUsed    ''=2*256=512
'
'        EConfig_Repeat_Cyc_for_Pgm = 1   ' 120
'        EConfigHiLimitSingleDoubleBitCheck = 0
'
'    ElseIf (gS_EFuse_Orientation = "SingleUp") Then
'
'        EConfigBlock = 1
'        EConfigBitsPerRow = 32
'        EConfigReadBitWidth = 32
'        EConfigRowPerBlock = 32 '16
'        EConfigReadCycle = EConfigRowPerBlock
'
'        EConfigBitPerBlockUsed = EConfigRowPerBlock * EConfigBitsPerRow ''=16*32=512
'        EConfigTotalBitCount = EConfigBlock * EConfigBitPerBlockUsed    ''=1*512=512
'
'        EConfig_Repeat_Cyc_for_Pgm = 1  '120
'        EConfigHiLimitSingleDoubleBitCheck = 0
'
'    ''''it's reserved.
'    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
'    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
'    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
'    End If
    
    If (gS_EFuse_Orientation = "SingleUp") Then
        m_divider = 32
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT" Or gS_EFuse_Orientation = "UP2DOWN") Then
        m_divider = 16
    End If
    
    'default is all zero here
    ReDim gL_CFG_SegFlag_arr(EConfigBitPerBlockUsed / m_divider - 1)

    gS_CFG_CRC_Stage = "" ''''initial
    m_condCateCNT = 0     ''''initial
    
    ''''find out the "firstbits" index, first 0~63 bits in general
    ''''Determin if using BinCut for the Vdd Binning by EFuse Revision
    ''''<Important> Using 'bincut' in the column 'Default or Real' to decide if VddBinning Fuse
    For i = 0 To UBound(CFGFuse.Category)
        m_algorithm = LCase(CFGFuse.Category(i).algorithm)
        m_defval = CFGFuse.Category(i).DefaultValue
        m_catename = CFGFuse.Category(i).Name
        m_stage = LCase(CFGFuse.Category(i).Stage)
        m_defreal = LCase(CFGFuse.Category(i).Default_Real)
        m_bitwidth = CFGFuse.Category(i).BitWidth
        m_LSBbit = CFGFuse.Category(i).LSBbit
        m_MSBBit = CFGFuse.Category(i).MSBbit

        If (m_algorithm = "crc") Then m_defval = 0

        ''''20171211, Here it's used to judge default value if is Hex or Binary
        m_defval = auto_checkDefaultValue(m_defval, m_algorithm, m_binarr, m_bitwidth, m_defreal)
        'm_defval = auto_checkDefaultValue(m_defval, m_binarr, m_bitwidth, m_defreal)
        CFGFuse.Category(i).DefValBitArr = m_binarr
        ms_defval = m_defval ''''201811XX, to SiteVariant
        
        If (gS_JobName = m_stage And m_defreal <> "default" And m_algorithm <> "cond") Then
            m_idx0 = Floor(m_LSBbit / m_divider) 'please check floor
            m_idx1 = Floor(m_MSBBit / m_divider) 'please check floor
            gL_CFG_SegFlag_arr(m_idx0) = 1
            If (m_idx0 <> m_idx1) Then
                For j = 1 To m_idx1 - m_idx0
                    gL_CFG_SegFlag_arr(m_idx0 + j) = 1
                Next
            End If
        End If
        
'        If (m_algorithm = "firstbits") Then
'            gI_CFG_firstbits_index = i
'            gS_CFG_firstbits_stage = m_stage
'
'        ElseIf (m_algorithm = "cond") Then ''''20170630 update
        If (m_algorithm = "cond") Then
            ''''------------------------------------------------------------------------------
            ''''here just do once
            If (gB_findCFGCondTable_flag And m_condCateCNT = 0) Then
                m_cmpValue_sum = 0
                m_condCateCNT = m_condCateCNT + 1
                gI_CFG_firstbits_index = i
                gS_CFG_firstbits_stage = "ft3" ''''default, was 'cp1', 20170804 update
                For j = 0 To UBound(CFGTable.Category(0).condition)
                    m_stage = LCase(CFGTable.Category(0).condition(j).Stage)

                    ''''20170804 update, to support other possibility except for CP1, FT3
                    If (gS_JobName <> gS_CFG_firstbits_stage And gS_JobName = m_stage) Then
                        gS_CFG_firstbits_stage = m_stage
                        Exit For
                    End If
                Next j
                
                gL_CFG_Cond_JobvsStage = -1 ''''default
                For j = 0 To UBound(CFGTable.Category(0).condition)
                    m_stage = LCase(CFGTable.Category(0).condition(j).Stage)

                    ''''check if current Job is less(-1)/equal(0)/large(1) than the stage of all condition bits
                    ''''gL_CFG_Cond_JobvsStage = 0 means: current Job is exited in all Stages of Cond
                    ''''gL_CFG_Cond_JobvsStage = 1 means: all Cond bits should have been fused
                    ''''gL_CFG_Cond_JobvsStage =-1 means: all Cond bits should have NOT been fused
                    m_cmpValue = auto_eFuse_check_Job_cmpare_Stage(m_stage)
                    m_cmpValue_sum = m_cmpValue_sum + m_cmpValue
                    If (j = 0) Then
                        gL_CFG_Cond_JobvsStage = m_cmpValue
                    ElseIf (m_cmpValue = 0) Then
                        gL_CFG_Cond_JobvsStage = 0
                        Exit For
                    End If
                Next j

                If (Abs(m_cmpValue_sum) = (1 + UBound(CFGTable.Category(0).condition))) Then
                    If (m_cmpValue_sum > 0) Then gL_CFG_Cond_JobvsStage = 1
                    If (m_cmpValue_sum = 0) Then gL_CFG_Cond_JobvsStage = 0
                    If (m_cmpValue_sum < 0) Then gL_CFG_Cond_JobvsStage = -1
                End If
            End If
            ''''------------------------------------------------------------------------------
            
            ''''201811XX
            If (gB_eFuse_newMethod = True) Then
                ''''New Method
                ''''20180918 New Method Start, Ref from MC2T---------------------
                If (kk <> -1) Then
                    m_condIdx = CFGCondTabIndex(m_catename)
                    With CFGFuse.Category(i)
                        .Default_Real = "Default"
                        .Stage = CFGTable.Category(kk).condition(m_condIdx).Stage
                        If (LCase(.Stage) = "cp1") Then .Stage = "CP1_EARLY" ''''<NOTICE>
                        .DefaultValue = CFGTable.Category(kk).condition(m_condIdx).Decimal
                        .PatTestPass_Flag = True
                        m_defval = .DefaultValue
                        m_stage = LCase(.Stage)
                        m_defreal = LCase(.Default_Real)
                        .DefValBitArr = CFGTable.Category(kk).condition(m_condIdx).BitVal ''''<MUST>
                    End With
                Else
                    ''''case Unknown
                    With CFGFuse.Category(i)
                        .Default_Real = "Default"
                        .Stage = "CP1_EARLY" ''''<NOTICE>
                        .DefaultValue = 0
                        .PatTestPass_Flag = True
                        m_stage = LCase(.Stage)
                        m_defreal = LCase(.Default_Real)
                    End With
                        ''''unused, check it later
''''                    If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0 Or gB_eFuse_CFG_Cond_FTF_done_Flag = True) Then
''''                        m_bitStrM = auto_Dec2Bin_EFuse(m_defval, m_bitwidth, m_binarr)
''''                    Else
''''                        ''''Job_cmpare_Stage =-1 means: cond bits should have NOT been fused
''''                        m_bitStrM = auto_Dec2Bin_EFuse(0, m_bitwidth, m_binarr)
''''                    End If
''''                    gS_CFGCondTable_bitsStr = m_bitStrM + gS_CFGCondTable_bitsStr
                End If
                ms_defval = m_defval
                ''''20180918 New Method End, ''''201811XX ---------------------
            End If
        ElseIf (m_algorithm = "revision") Then
            If (m_defreal = "real") Then
                Call auto_eFuse_SetWriteVariable_SiteAware("CFG", m_catename, ms_defval, False)
            End If
'        ElseIf (m_algorithm = "fuse" Or m_algorithm = "device" Or m_algorithm = "revision") Then
'            If (m_defreal = "real") Then
'                Call auto_eFuse_SetWriteVariable_SiteAware("CFG", m_catename, ms_defval, False)
'''                For Each ss In TheExec.Sites
'''                    Call auto_eFuse_SetWriteDecimal("CFG", m_catename, m_defval, False)
'''                Next ss
'            End If

        ElseIf (m_algorithm = "base") Then ''''M8 case
            m_resolution = CFGFuse.Category(i).Resoultion
            If (m_defreal = "decimal") Then
                gD_VBaseFuse = m_defval
                If (m_resolution = 0) Then
                    gD_BaseStepVoltage = 25#
                Else
                    gD_BaseStepVoltage = m_resolution
                End If
                gD_BaseVoltage = (gD_VBaseFuse + 1) * gD_BaseStepVoltage
            Else ''''m_defreal = "default" or "safe voltage"
                ''''Base Fuse = 400mV = 400/25 -1 =15 = "01111"
                gD_BaseStepVoltage = m_resolution
                gD_BaseVoltage = CFGFuse.Category(i).DefaultValue         ''If  Current base voltage = 400
                gD_VBaseFuse = (gD_BaseVoltage / gD_BaseStepVoltage) - 1  ''Then After calculating, it's 15
            End If
        
        ''''' 20161003 ADD CRC
        ElseIf (m_algorithm = "crc") Then
            gS_CFG_CRC_Stage = m_stage
            ''''20161026 add
            gL_CFG_CRC_LSBbit = CFGFuse.Category(i).LSBbit
            gL_CFG_CRC_MSBbit = CFGFuse.Category(i).MSBbit
            gL_CFG_CRC_BitWidth = CFGFuse.Category(i).BitWidth
            If ((Abs(gL_CFG_CRC_MSBbit - gL_CFG_CRC_LSBbit + 1) - gL_CFG_CRC_BitWidth) <> 0) Then
                TheExec.AddOutput funcName + ":: Bitwidth is NOT equal to (MSBbit-LSBbit)"
                GoTo errHandler
            End If

            ''''----------------------------------------------------------------------------------------
            ''''Example
            ''''----------------------------------------------------------------------------------------
            ''''Comment/Description:: CRC_IgnoreBits=[287:0],[387:380],[403,396],[511:496]
            ''''----------------------------------------------------------------------------------------
            m_cmt = Trim(UCase(CFGFuse.Category(i).comment))
            m_cmtArr = Split(m_cmt, ",")
            ReDim gL_CFG_CRC_calcBits(EConfigBitPerBlockUsed - 1) ''''<MUST>be here
            ''''<MUST>initial all bits = 1 as default
            For j = 0 To UBound(gL_CFG_CRC_calcBits)
                gL_CFG_CRC_calcBits(j) = 1
            Next j

            m_find_CFG_calcBits_flag = False ''''default

            For j = 0 To UBound(m_cmtArr)
                ''''reuse the variable
                m_cmt = Trim(m_cmtArr(j))
                m_cmt = Replace(m_cmt, "[", "", 1)
                m_cmt = Replace(m_cmt, "]", "", 1)
                m_bit_min = 99999
                m_bit_max = -99999

                If (m_cmt Like UCase("*ignoreBits*=*[*:*]*")) Then ''''<MUST>
                    m_find_CFG_calcBits_flag = True
                End If
                
                If (m_cmt Like UCase("*=*")) Then
                    m_tmpArr = Split(m_cmt, "=")
                    m_cmt = Trim(m_tmpArr(1))
                End If
                m_tmpArr = Split(m_cmt, ":")
                For k = 0 To UBound(m_tmpArr)
                    m_cmt = Trim(m_tmpArr(k))
                    If (IsNumeric(m_cmt)) Then
                        m_value = CLng(m_cmt)
                        If (m_value > m_bit_max) Then
                            m_bit_max = m_value
                        Else
                            m_bit_min = m_value
                        End If
                    Else
                        TheExec.AddOutput funcName + ":: bits number is not a Numeric (" + m_cmt + ")"
                        GoTo errHandler
                    End If
                Next k
                ''''these bits MUST be excluded.
                For n = m_bit_min To m_bit_max
                    gL_CFG_CRC_calcBits(n) = 0
                Next n
            Next j

            If (m_find_CFG_calcBits_flag = False) Then
                ''''Exception Case
                TheExec.AddOutput funcName + ":: Problem on CFG CRC Comment Format, no calcBits definition, " + CFGFuse.Category(i).comment
                TheExec.Datalog.WriteComment funcName + ":: Problem on CFG CRC Comment Format, no calcBits definition, " + CFGFuse.Category(i).comment
                GoTo errHandler ''''<MUST>
            End If

            ''''<MUST> calcBits MUST exclude CFG_CRC selfbits
            ''''<NOTICE> Here is used to prevent any exception once the comment/description does not include these bits
            m_bit_min = CFGFuse.Category(i).LSBbit
            m_bit_max = CFGFuse.Category(i).MSBbit
            If (m_bit_max < m_bit_min) Then
                m_bit_max = CFGFuse.Category(i).LSBbit
                m_bit_min = CFGFuse.Category(i).MSBbit
            End If
            For n = m_bit_min To m_bit_max
                gL_CFG_CRC_calcBits(n) = 0
            Next n
            
        ''''It's the condition of Tgib below.
''''            If (UCase(m_catename) = "CFG_CRC") And (LCase(gS_JobName) = "ft1" Or LCase(gS_JobName) = "ft2") Then  '''''ken 20161006
''''                CFGFuse.Category(i).Stage = LCase(gS_JobName)
''''                gS_CFG_CRC_Stage = LCase(gS_JobName)
''''            End If
            
        ''''It's the condition of Tgib below.
''''        ''''' These bit always need to be fused but we need fused it at FT2 stage in bringup and fused it at FT1 stage in production
''''        ElseIf (m_algorithm = "fcal") Then
''''            If (UCase(m_catename) = "PCIE_REFPLL_FCAL_VCO_DIGCTRL") And (LCase(gS_JobName) = "ft1" Or LCase(gS_JobName) = "ft2") Then  '''''ken 20160811
''''                CFGFuse.Category(i).Stage = LCase(gS_JobName)
''''            End If
        
        ''''20170630 update
        ''ElseIf (m_algorithm = "scan") Then ''''was
        ElseIf (m_algorithm = "scan" And gS_JobName Like "*cp1*") Then
            gS_CFG_SCAN_stage = m_stage ''''need to check ?? 201811XX
        End If

        ''''20160328 update, 20160624 update
        If (m_defreal <> "real" And m_defreal <> "bincut") Then
            CFGFuse.Category(i).DefaultValue = m_defval
            CFGFuse.Category(i).HiLMT = m_defval
            CFGFuse.Category(i).LoLMT = m_defval
            If (m_algorithm = "base") Then ''''M8 case <MUST be here>
                 If (m_defreal = "default") Then
                'If (m_defreal Like "safe*voltage") Then
                    m_defval = gD_VBaseFuse
                    ms_defval = gD_VBaseFuse
                End If
            End If
            ''''20180712, will do the update later on for "vddbin: safe voltage"
            ''''''''''''' please see auto_precheck_SafeVoltage_Base_VddBin()
'            If (m_algorithm = "vddbin" And m_defreal Like "safe*voltage") Then
'            Else
''                For Each ss In TheExec.Sites.Existing
''                    Call auto_eFuse_SetWriteDecimal("CFG", m_catename, m_defval, False)
''                Next ss
'                Call auto_eFuse_SetWriteVariable_SiteAware("CFG", m_catename, ms_defval, False)
'            End If
            ''''20180723, For "Real", give an initial value as m_defval, usually it is "Zero".
'            For Each ss In TheExec.Sites.Existing
'                Call auto_eFuse_SetWriteDecimal("CFG", m_catename, m_defval, False)
'            Next ss
        End If
        
        If (m_algorithm = "vddbin" And m_defreal = "default") Then
        'If (m_algorithm = "vddbin" And m_defreal Like "safe*voltage") Then
        Else
            Call auto_eFuse_SetWriteVariable_SiteAware("CFG", m_catename, ms_defval, False)
        End If
    Next i
    
    gL_CFG_SegCNT = 0
    For i = 0 To UBound(gL_CFG_SegFlag_arr)
        If (gL_CFG_SegFlag_arr(i) = 1) Then gL_CFG_SegCNT = gL_CFG_SegCNT + 1
    Next i

    'Set eFuse Global Data initial
    For Each ss In TheExec.sites.Existing
        gB_CFGSVM_BIT_Read_ValueisONE(ss) = False ''''<MUST>
        gS_CFG_CRC_HexStr(ss) = "" ''''MUST be
    Next ss

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_UIDConstant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UIDConstant_Initialize"
    
    Dim i As Long
    Dim idx As Long
    Dim ss As Variant
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_bitwidth As Long
    Dim m_defval As Variant
    Dim m_len As Long
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_defreal As String
    Dim m_defvalhex As String ''''without prefix "0x"
    Dim m_binarr() As Long

    DisplayUID = True ''''<User Maintain>

    ''''Get max length of category name
    gI_UID_catename_maxLen = 0
    gL_UIDCodeBitWidth = 0 ''''initial
    For i = 0 To UBound(UIDFuse.Category)
        m_catename = UIDFuse.Category(i).Name
        m_algorithm = LCase(UIDFuse.Category(i).algorithm)
        m_defval = UIDFuse.Category(i).DefaultValue
        m_bitwidth = UIDFuse.Category(i).BitWidth
        m_len = Len(UIDFuse.Category(i).Name)
        m_lolmt = UIDFuse.Category(i).LoLMT
        m_hilmt = UIDFuse.Category(i).HiLMT
        m_defreal = LCase(UIDFuse.Category(i).Default_Real)

        ''''20171211, Here it's used to judge default value if is Hex or Binary
        m_defval = auto_checkDefaultValue(m_defval, m_algorithm, m_binarr, m_bitwidth, m_defreal)

        'm_defval = auto_checkDefaultValue(m_defval, m_binarr, m_bitwidth, m_defreal)
        UIDFuse.Category(i).DefValBitArr = m_binarr

        If (m_len > gI_UID_catename_maxLen) Then
            gI_UID_catename_maxLen = m_len
        End If

        If (m_algorithm = "uid") Then
            gL_UIDCodeBitWidth = gL_UIDCodeBitWidth + m_bitwidth ''''because multiple 'uid' categories
            UID_ChkSum_LoLimit = m_lolmt
            UID_ChkSum_HiLimit = m_hilmt
        
        ElseIf (m_algorithm = "crc") Then
            ''''doNothing
            gS_UID_CRC_Stage = LCase(UIDFuse.Category(i).Stage)
            gL_UID_CRC_EndBit = UIDFuse.Category(i).LSBbit - 1
            gL_UID_CRC_LSBbit = UIDFuse.Category(i).LSBbit
            gL_UID_CRC_MSBbit = UIDFuse.Category(i).MSBbit
            gL_UID_CRC_BitWidth = m_bitwidth
        ElseIf (m_algorithm = "rid") Then
            For Each ss In TheExec.sites.Existing
                Call auto_eFuse_SetWriteDecimal("UID", m_catename, m_defval, False)
            Next ss
        End If

        ''''20160328 update, 20160624 update
        If (m_defreal <> "real" And m_defreal <> "bincut") Then
            UIDFuse.Category(i).DefaultValue = m_defval
            UIDFuse.Category(i).HiLMT = m_defval
            UIDFuse.Category(i).LoLMT = m_defval
            For Each ss In TheExec.sites.Existing
                Call auto_eFuse_SetWriteDecimal("UID", m_catename, m_defval, False)
            Next ss
        End If
    Next i
    gI_UID_catename_maxLen = gI_UID_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------
    
    ''''<Important>
    UIDBitsPerCode = 128 ''''Because there are 128 random bits are generated from C651 *.dll file
    gL_UIDCode_Block = gL_UIDCodeBitWidth / UIDBitsPerCode
    
    If (gS_EFuse_Orientation = "UP2DOWN") Then

        UIDBlock = 2
        UIDRowPerBlock = 32
        UIDBitsPerRow = 32
        UIDWriteBitExpandWidth = 120
        UIDReadBitWidth = 32
        
        UIDBitsPerBlockUsed = UIDRowPerBlock * UIDBitsPerRow ''=32x32=1024
        UIDBitsPerBlock = UIDBitsPerBlockUsed                ''=1024
        UIDReadCycle = UIDRowPerBlock * UIDBlock             ''=32*2=64 <Notice>
        UIDTotalBits = UIDBitsPerBlockUsed * UIDBlock        ''=1024x2=2048
        UIDBitsPerCycle = UIDReadBitWidth                    ''=32

    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then

        UIDBlock = 2
        UIDRowPerBlock = 32
        UIDBitsPerRow = 16
        UIDWriteBitExpandWidth = 120
        UIDReadBitWidth = 32
        
        UIDBitsPerBlockUsed = UIDRowPerBlock * UIDBitsPerRow ''=64x16=1024  , ''=32x16=512
        UIDBitsPerBlock = UIDBitsPerBlockUsed                ''=1024        , ''=512
        UIDReadCycle = UIDRowPerBlock                        ''=64 <Notice> , ''=32
        UIDTotalBits = UIDBitsPerBlockUsed * UIDBlock        ''=1024x2=2048 , ''=512x2=1024
        UIDBitsPerCycle = UIDReadBitWidth                    ''=32          , ''=32

    ElseIf (gS_EFuse_Orientation = "SingleUp") Then
        
        UIDBlock = 1          ''<Important>
        UIDRowPerBlock = 32
        UIDBitsPerRow = 32
        UIDWriteBitExpandWidth = 120
        UIDReadBitWidth = 32
        
        UIDBitsPerBlockUsed = UIDRowPerBlock * UIDBitsPerRow ''=32x32=1024
        UIDBitsPerBlock = UIDBitsPerBlockUsed                ''=1024
        UIDReadCycle = UIDRowPerBlock * UIDBlock             ''=32*1=32 <Notice>
        UIDTotalBits = UIDBitsPerBlockUsed * UIDBlock        ''=1024x1=1024
        UIDBitsPerCycle = UIDReadBitWidth                    ''=32
     
    ''''Below is reserved for the future, there is NO any definition at present.
    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
    End If
    
    Dim mDL_UID_CRC_Size As New DSPWave
            
    gDW_UID_CRC_calcBits_Temp.CreateConstant 1, UIDBitsPerBlock, DspLong
    mDL_UID_CRC_Size.CreateConstant 0, UIDBitsPerBlock - gL_UID_CRC_EndBit, DspLong
    gDW_UID_CRC_calcBits_Temp.Select(gL_UID_CRC_EndBit + 1, 1, UIDBitsPerBlock).Replace mDL_UID_CRC_Size


Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_UDRConstant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRConstant_Initialize"

    Dim i As Long
    Dim ss As Variant
    Dim m_algorithm As String
    Dim m_catename As String
    Dim m_resolution As Double
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_usiLSBcycle As Long
    Dim m_usiMSBcycle As Long
    Dim m_usoLSBcycle As Long
    Dim m_usoMSBcycle As Long
    Dim m_len As Long
    Dim m_defval As Variant
    Dim m_usipretmp As String
    Dim m_usopretmp As String
    Dim m_defreal As String
    Dim m_bitwidth As Long
    Dim m_defvalhex As String ''''without prefix "0x"
    Dim m_binarr() As Long
    Dim ms_defval As New SiteVariant

    ''''Get max length of category name
    gI_UDR_catename_maxLen = 0
    For i = 0 To UBound(UDRFuse.Category)
        m_len = Len(UDRFuse.Category(i).Name)
        m_MSBBit = UDRFuse.Category(i).MSBbit

        If (m_len > gI_UDR_catename_maxLen) Then
            gI_UDR_catename_maxLen = m_len
        End If
    Next i
    gI_UDR_catename_maxLen = gI_UDR_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------

    m_usipretmp = ""
    m_usopretmp = ""

    For i = 0 To UBound(UDRFuse.Category)
        m_catename = UDRFuse.Category(i).Name
        m_algorithm = LCase(UDRFuse.Category(i).algorithm)
        m_defval = UDRFuse.Category(i).DefaultValue
        m_resolution = UDRFuse.Category(i).Resoultion
        m_LSBbit = UDRFuse.Category(i).LSBbit
        m_MSBBit = UDRFuse.Category(i).MSBbit
        m_usiLSBcycle = UDRFuse.Category(i).USILSBCycle
        m_usiMSBcycle = UDRFuse.Category(i).USIMSBCycle
        m_usoLSBcycle = UDRFuse.Category(i).USOLSBCycle
        m_usoMSBcycle = UDRFuse.Category(i).USOMSBCycle
        m_defreal = LCase(UDRFuse.Category(i).Default_Real)
        m_bitwidth = UDRFuse.Category(i).BitWidth

        ''''-------------------------------------------------------
        If (m_usiLSBcycle < m_usiMSBcycle) Then
            gL_USI_PatBitOrder = "LSB"
        ElseIf (m_usiLSBcycle > m_usiMSBcycle) Then
            gL_USI_PatBitOrder = "MSB"
        End If
        If (m_usipretmp <> "" And gL_USI_PatBitOrder <> m_usipretmp) Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: " + m_catename + " the USI bit order is different from others"
            GoTo errHandler
        End If
        m_usipretmp = gL_USI_PatBitOrder
        ''''-------------------------------------------------------
        
        ''''-------------------------------------------------------
        If (m_usoLSBcycle < m_usoMSBcycle) Then
            gL_USO_PatBitOrder = "LSB"
        ElseIf (m_usoLSBcycle > m_usoMSBcycle) Then
            gL_USO_PatBitOrder = "MSB"
        End If
        If (m_usopretmp <> "" And gL_USO_PatBitOrder <> m_usopretmp) Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: " + m_catename + " the USO bit order is different from others"
            GoTo errHandler
        End If
        m_usopretmp = gL_USO_PatBitOrder
        ''''-------------------------------------------------------
        
        ''''20171211, Here it's used to judge default value if is Hex or Binary
        m_defval = auto_checkDefaultValue(m_defval, m_algorithm, m_binarr, m_bitwidth, m_defreal)
        'm_defval = auto_checkDefaultValue(m_defval, m_binarr, m_bitwidth, m_defreal)
        UDRFuse.Category(i).DefValBitArr = m_binarr
        ms_defval = m_defval

        If (m_algorithm = "base") Then
            If (m_defreal = "decimal") Then ''''20160624 update
                gD_VBaseFuse = m_defval
                If (m_resolution = 0) Then
                    gD_BaseStepVoltage = 25#
                Else
                    gD_BaseStepVoltage = m_resolution
                End If
                gD_BaseVoltage = (gD_VBaseFuse + 1) * gD_BaseStepVoltage
            Else
                gD_BaseVoltage = m_defval
                gD_BaseStepVoltage = m_resolution
                gD_VBaseFuse = (gD_BaseVoltage / gD_BaseStepVoltage) - 1
            End If

        ElseIf (m_algorithm = "fuse") Then
            gL_UDR_eFuse_Revision = m_defval
            For Each ss In TheExec.sites.Existing
                Call auto_eFuse_SetWriteDecimal("UDR", m_catename, m_defval, False)
            Next ss
        
        ''ElseIf (m_algorithm = "lastreserve") Then
        ElseIf (i = UBound(UDRFuse.Category)) Then
            'gL_USI_DigSrcBits_Num
            If (m_LSBbit <= m_MSBBit) Then
                gL_USI_DigSrcBits_Num = m_MSBBit + 1
                gL_USO_DigCapBits_Num = m_MSBBit + 1
            Else
                gL_USI_DigSrcBits_Num = m_LSBbit + 1
                gL_USO_DigCapBits_Num = m_LSBbit + 1
            End If
        End If

        ''''20160328 update, 20160624 update
        If (m_defreal <> "real" And m_defreal <> "bincut") Then
            UDRFuse.Category(i).DefaultValue = m_defval
            UDRFuse.Category(i).HiLMT = m_defval
            UDRFuse.Category(i).LoLMT = m_defval
            If (m_algorithm = "base") Then
                If (m_defreal = "default") Then
                'If (m_defreal Like "safe*voltage") Then
                    m_defval = gD_VBaseFuse
                    ms_defval = gD_VBaseFuse
                End If
            End If

            ''''20180712, will do the update later on for "vddbin: safe voltage"
            ''''''''''''' please see auto_precheck_SafeVoltage_Base_VddBin()

        End If
        'If (m_algorithm = "vddbin" And m_defreal <> "decimal") Then
        'If (m_algorithm = "vddbin" And m_defreal Like "safe*voltage") Then
        If (m_algorithm = "vddbin" And m_defreal = "default") Then
        Else
'                For Each ss In TheExec.Sites.Existing
'                    Call auto_eFuse_SetWriteDecimal("UDR", m_catename, m_defval, False)
'                Next ss
            Call auto_eFuse_SetWriteVariable_SiteAware("UDR", m_catename, ms_defval, False)
        End If
    Next i

    'Set eFuse Global Data initial
    For Each ss In TheExec.sites.Existing
        gS_USI_BitStr(ss) = ""
    Next ss
    
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment funcName + "::"
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment FormatNumeric("USI_PatBitOrder = ", 25) + FormatNumeric(gL_USI_PatBitOrder, -10)
    TheExec.Datalog.WriteComment FormatNumeric("USI_DigSrcBits_Num = ", 25) + FormatNumeric(gL_USI_DigSrcBits_Num, -10)
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment FormatNumeric("USO_PatBitOrder = ", 25) + FormatNumeric(gL_USO_PatBitOrder, -10)
    TheExec.Datalog.WriteComment FormatNumeric("USO_DigCapBits_Num = ", 25) + FormatNumeric(gL_USO_DigCapBits_Num, -10)
    TheExec.Datalog.WriteComment "---------------------------------------------"
    ''TheExec.Datalog.WriteComment ""

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_SENConstant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_SENConstant_Initialize"

    Dim i As Long
    Dim ss As Variant
    Dim m_algorithm As String
    Dim m_MSBBit As Long
    Dim m_LSBbit As Long
    Dim m_len As Long
    Dim m_catename As String
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim m_bitwidth As Long
    Dim m_defvalhex As String ''''without prefix "0x"
    Dim m_binarr() As Long

    ''''Get max length of category name
    gI_SEN_catename_maxLen = 0
    For i = 0 To UBound(SENFuse.Category)
        m_len = Len(SENFuse.Category(i).Name)
        m_MSBBit = SENFuse.Category(i).MSBbit
        
        If (m_len > gI_SEN_catename_maxLen) Then
            gI_SEN_catename_maxLen = m_len
        End If
    Next i
    gI_SEN_catename_maxLen = gI_SEN_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------

    If (gS_EFuse_Orientation = "UP2DOWN") Then

        SENSORBlock = 2
        SENSORRowPerBlock = 16
        SENSORReadBitWidth = 32
        SENSORBitsPerRow = 32
        SENSORReadCycle = SENSORRowPerBlock * SENSORBlock

        SENSOR_Repeat_Cyc_for_Pgm = 120
        SENSORBitPerBlockUsed = SENSORBitsPerRow * SENSORRowPerBlock ''''=32*16=512
        SENSORTotalBitCount = SENSORBitPerBlockUsed * SENSORBlock    ''''=512*2=1024
        SENSORHiLimitSingleDoubleBitCheck = 0

    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then

        SENSORBlock = 2
        SENSORRowPerBlock = 16
        SENSORReadBitWidth = 32
        SENSORBitsPerRow = 16            ''''here must be 16
        SENSORReadCycle = SENSORRowPerBlock
        
        SENSOR_Repeat_Cyc_for_Pgm = 120
        SENSORBitPerBlockUsed = SENSORBitsPerRow * SENSORRowPerBlock ''''=16*32=512   ,=16*16=256
        SENSORTotalBitCount = SENSORBitPerBlockUsed * SENSORBlock    ''''=512*2=1024  ,=256*2=512
        SENSORHiLimitSingleDoubleBitCheck = 0

    ElseIf (gS_EFuse_Orientation = "SingleUp") Then

        SENSORBlock = 1
        SENSORRowPerBlock = 16
        SENSORReadBitWidth = 32
        SENSORBitsPerRow = 32
        SENSORReadCycle = SENSORRowPerBlock

        SENSOR_Repeat_Cyc_for_Pgm = 120
        SENSORBitPerBlockUsed = SENSORBitsPerRow * SENSORRowPerBlock ''''=32*16=512
        SENSORTotalBitCount = SENSORBitPerBlockUsed * SENSORBlock    ''''=512*1=512
        SENSORHiLimitSingleDoubleBitCheck = 0
    
    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
    End If

    For i = 0 To UBound(SENFuse.Category)
        m_algorithm = LCase(SENFuse.Category(i).algorithm)
        m_MSBBit = SENFuse.Category(i).MSBbit
        m_LSBbit = SENFuse.Category(i).LSBbit
        m_defval = SENFuse.Category(i).DefaultValue
        m_catename = SENFuse.Category(i).Name
        m_defreal = LCase(SENFuse.Category(i).Default_Real)
        m_bitwidth = SENFuse.Category(i).BitWidth

        ''''20171211, Here it's used to judge default value if is Hex or Binary
        m_defval = auto_checkDefaultValue(m_defval, m_algorithm, m_binarr, m_bitwidth, m_defreal)

        'm_defval = auto_checkDefaultValue(m_defval, m_binarr, m_bitwidth, m_defreal)
        SENFuse.Category(i).DefValBitArr = m_binarr

        If (m_algorithm = "crc") Then
            gS_SEN_CRC_Stage = LCase(SENFuse.Category(i).Stage)
            gL_SEN_CRC_EndBit = m_LSBbit - 1
            gL_SEN_CRC_LSBbit = SENFuse.Category(i).LSBbit
            gL_SEN_CRC_MSBbit = SENFuse.Category(i).MSBbit
            gL_SEN_CRC_BitWidth = SENFuse.Category(i).BitWidth
        End If

        ''''20160328 update, 20160624 update
        If (m_defreal <> "real" And m_defreal <> "bincut") Then
            SENFuse.Category(i).DefaultValue = m_defval
            SENFuse.Category(i).HiLMT = m_defval
            SENFuse.Category(i).LoLMT = m_defval
            For Each ss In TheExec.sites.Existing
                Call auto_eFuse_SetWriteDecimal("SEN", m_catename, m_defval, False)
            Next ss
        End If
    Next i
      
    '' for CRC case
    Dim mDL_SEN_CRC_Size As New DSPWave
            
    gDW_SEN_CRC_calcBits_Temp.CreateConstant 1, SENSORBitPerBlockUsed, DspLong
    mDL_SEN_CRC_Size.CreateConstant 0, SENSORBitPerBlockUsed - gL_SEN_CRC_EndBit, DspLong
    gDW_SEN_CRC_calcBits_Temp.Select(gL_SEN_CRC_EndBit + 1, 1, SENSORBitPerBlockUsed).Replace mDL_SEN_CRC_Size
    
    'Set eFuse Global Data initial
    For Each ss In TheExec.sites.Existing
        gS_SEN_CRC_HexStr(ss) = "" ''''MUST be
    Next ss

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_MONConstant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MONConstant_Initialize"

    Dim i As Long
    Dim ss As Variant
    Dim m_algorithm As String
    Dim m_MSBBit As Long
    Dim m_LSBbit As Long
    Dim m_len As Long
    Dim m_catename As String
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim m_bitwidth As Long
    Dim m_defvalhex As String ''''without prefix "0x"
    Dim m_binarr() As Long
    Dim ms_defval As New SiteVariant

    ''''Get max length of category name
    gI_MON_catename_maxLen = 0
    For i = 0 To UBound(MONFuse.Category)
        m_len = Len(MONFuse.Category(i).Name)
        m_MSBBit = MONFuse.Category(i).MSBbit
        
        If (m_len > gI_MON_catename_maxLen) Then
            gI_MON_catename_maxLen = m_len
        End If
    Next i
    gI_MON_catename_maxLen = gI_MON_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------
    
    ''20191230 , make bank size to be dynamic
    Dim m_Effective As Long
    m_Effective = MONFuse.Category(UBound(MONFuse.Category)).MSBbit + 1
    'MONITORReadBitWidth = 32
    MONITOR_Repeat_Cyc_for_Pgm = 1
    
    If (gS_EFuse_Orientation = "UP2DOWN") Then

        MONITORBlock = 2
        MONITORBitsPerRow = MONITORReadBitWidth
        MONITORRowPerBlock = m_Effective / MONITORReadBitWidth
        MONITORReadCycle = MONITORRowPerBlock * MONITORBlock

    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then

        MONITORBlock = 2
        MONITORBitsPerRow = MONITORReadBitWidth / MONITORBlock
        MONITORRowPerBlock = m_Effective / MONITORBitsPerRow
                   ''''here must be 16
        MONITORReadCycle = MONITORRowPerBlock
        
    ElseIf (gS_EFuse_Orientation = "SingleUp") Then

        MONITORBlock = 1
        MONITORBitsPerRow = MONITORReadBitWidth
        MONITORRowPerBlock = m_Effective / MONITORReadBitWidth
        
        MONITORReadCycle = MONITORRowPerBlock
    
    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
    End If
    
    MONITORBitPerBlockUsed = MONITORBitsPerRow * MONITORRowPerBlock ''''=32*16=512
    MONITORTotalBitCount = MONITORBitPerBlockUsed * MONITORBlock    ''''=512*2=1024

'    If (gS_EFuse_Orientation = "UP2DOWN") Then
'
'        MONITORBlock = 2
'        MONITORRowPerBlock = 16
'        MONITORReadBitWidth = 32
'        MONITORBitsPerRow = 32
'        MONITORReadCycle = MONITORRowPerBlock * MONITORBlock
'
'        MONITOR_Repeat_Cyc_for_Pgm = 120
'        MONITORBitPerBlockUsed = MONITORBitsPerRow * MONITORRowPerBlock ''''=32*16=512
'        MONITORTotalBitCount = MONITORBitPerBlockUsed * MONITORBlock    ''''=512*2=1024
'        MONITORHiLimitSingleDoubleBitCheck = 0
'
'    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
'
'        MONITORBlock = 2
'        MONITORRowPerBlock = 16
'        MONITORReadBitWidth = 32
'        MONITORBitsPerRow = 16            ''''here must be 16
'        MONITORReadCycle = MONITORRowPerBlock
'
'        MONITOR_Repeat_Cyc_for_Pgm = 1
'        MONITORBitPerBlockUsed = MONITORBitsPerRow * MONITORRowPerBlock ''''=16*32=512  , =16*16=256
'        MONITORTotalBitCount = MONITORBitPerBlockUsed * MONITORBlock    ''''=512*2=1024 , =256*2=512
'        MONITORHiLimitSingleDoubleBitCheck = 0
'
'    ElseIf (gS_EFuse_Orientation = "SingleUp") Then
'
'        MONITORBlock = 1
'        MONITORRowPerBlock = 16
'        MONITORReadBitWidth = 32
'        MONITORBitsPerRow = 32
'        MONITORReadCycle = MONITORRowPerBlock
'
'        MONITOR_Repeat_Cyc_for_Pgm = 120
'        MONITORBitPerBlockUsed = MONITORBitsPerRow * MONITORRowPerBlock ''''=32*16=512
'        MONITORTotalBitCount = MONITORBitPerBlockUsed * MONITORBlock    ''''=512*1=512
'        MONITORHiLimitSingleDoubleBitCheck = 0
'
'    ElseIf (gS_EFuse_Orientation = "SingleDown") Then
'    ElseIf (gS_EFuse_Orientation = "SingleRight") Then
'    ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
'    End If

    For i = 0 To UBound(MONFuse.Category)
        m_algorithm = LCase(MONFuse.Category(i).algorithm)
        m_MSBBit = MONFuse.Category(i).MSBbit
        m_LSBbit = MONFuse.Category(i).LSBbit
        m_defval = MONFuse.Category(i).DefaultValue
        m_catename = MONFuse.Category(i).Name
        m_defreal = LCase(MONFuse.Category(i).Default_Real)
        m_bitwidth = MONFuse.Category(i).BitWidth

        ''''20171211, Here it's used to judge default value if is Hex or Binary
        m_defval = auto_checkDefaultValue(m_defval, m_algorithm, m_binarr, m_bitwidth, m_defreal)

        'm_defval = auto_checkDefaultValue(m_defval, m_binarr, m_bitwidth, m_defreal)
        MONFuse.Category(i).DefValBitArr = m_binarr
        ms_defval = m_defval

        If (m_algorithm = "crc") Then
            gL_MON_CRC_LSBbit = MONFuse.Category(i).LSBbit
            gL_MON_CRC_MSBbit = MONFuse.Category(i).MSBbit
            gL_MON_CRC_BitWidth = MONFuse.Category(i).BitWidth
            gS_MON_CRC_Stage = LCase(MONFuse.Category(i).Stage)
            gL_MON_CRC_EndBit = m_LSBbit - 1
        End If
        ''''20160328 update, 20160624 update
        If (m_defreal <> "real" And m_defreal <> "bincut") Then
            MONFuse.Category(i).DefaultValue = m_defval
            MONFuse.Category(i).HiLMT = m_defval
            MONFuse.Category(i).LoLMT = m_defval
'            For Each ss In TheExec.Sites.Existing
'                Call auto_eFuse_SetWriteDecimal("MON", m_catename, m_defval, False)
'            Next ss
            'Call auto_eFuse_SetWriteVariable_SiteAware("MON", m_catename, ms_defval, False)
        End If
        If (m_algorithm = "vddbin" And m_defreal <> "decimal") Then
        Else
'                For Each ss In TheExec.Sites.Existing
'                    Call auto_eFuse_SetWriteDecimal("UDR", m_catename, m_defval, False)
'                Next ss
            Call auto_eFuse_SetWriteVariable_SiteAware("MON", m_catename, ms_defval, False)
        End If
    Next i
    
        '' for CRC case
    'Dim mDL_MON_CRC_Calc As New DSPWave
    Dim mDL_MON_CRC_Size As New DSPWave
            
    gDW_MON_CRC_calcBits_Temp.CreateConstant 1, MONITORBitPerBlockUsed, DspLong
    mDL_MON_CRC_Size.CreateConstant 0, MONITORBitPerBlockUsed - gL_MON_CRC_EndBit, DspLong
    gDW_MON_CRC_calcBits_Temp.Select(gL_MON_CRC_EndBit + 1, 1, MONITORBitPerBlockUsed).Replace mDL_MON_CRC_Size

    
    'Set eFuse Global Data initial
    For Each ss In TheExec.sites.Existing
        gS_MON_CRC_HexStr(ss) = "" ''''MUST be
    Next ss

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160608 Add, 20180725 update
''''It's used to pre-check if the baseVoltage and Vddbin safeVoltage (Default Value) is multiple of the resolution
Public Function auto_precheck_SafeVoltage_Base_VddBin()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_precheck_SafeVoltage_Base_VddBin"

    Dim i As Long
    Dim ss As Variant
    Dim m_algorithm As String
    Dim m_catename As String
    Dim m_defval As Variant
    Dim m_tmpVal As Double
    Dim m_delta As Double
    Dim m_len As Long
    Dim m_resolution As Double
    Dim m_defreal As String
    Dim m_bitwidth As Long
    Dim m_binarr() As Long
    Dim m_defval_calc As Variant
    Dim m_updateFlag As Boolean

    ''''CFG Fuse
    If (gB_findCFG_flag) Then
        For i = 0 To UBound(CFGFuse.Category)
            m_algorithm = LCase(CFGFuse.Category(i).algorithm)
            If (m_algorithm = "base" Or m_algorithm = "vddbin") Then
                m_defreal = LCase(CFGFuse.Category(i).Default_Real)
                m_catename = CFGFuse.Category(i).Name
                m_resolution = CFGFuse.Category(i).Resoultion
                m_defval = CFGFuse.Category(i).DefaultValue
                m_bitwidth = CFGFuse.Category(i).BitWidth
                m_defval_calc = 0 ''''<MUST>
                m_updateFlag = False

                 If (m_defreal = "default") Then
                'If (m_defreal Like "safe*voltage") Then
                    If (m_algorithm = "base") Then
                        m_tmpVal = m_defval / m_resolution - 1 ''''<NOTICE>
                    ElseIf (m_algorithm = "vddbin") Then
                        m_tmpVal = (m_defval - gD_BaseVoltage) / m_resolution
                    End If
                    m_delta = Abs(m_tmpVal - Int(m_tmpVal))
                    If (m_delta > 0) Then
                        m_updateFlag = False
                        TheExec.Datalog.WriteComment "<Error> " + m_catename + ":: Default Value (" & m_defval & ") is Not multiple of the resolution (" & m_resolution & ")"
                        GoTo errHandler
                    Else
                        m_updateFlag = True
                        ''''m_delta=0, here using "default" to get the binarr with the calculated value
                        m_defval_calc = auto_checkDefaultValue(m_tmpVal, m_algorithm, m_binarr, m_bitwidth, "safe voltage")
                        'm_defval_calc = auto_checkDefaultValue(m_tmpVal, m_binarr, m_bitwidth, "default")
                        CFGFuse.Category(i).DefValBitArr = m_binarr
                    End If
                Else
                    ''''other cases (bincut,default,decimal)
                    ''''do Nothing
                    ''''It has been done in the auto_CFGConstant_Initialize()
                End If
                
                If (m_updateFlag = True) Then
                    ''''Update the correct default value to the Write structure
                    For Each ss In TheExec.sites.Existing
                        Call auto_eFuse_SetWriteDecimal("CFG", m_catename, m_defval_calc, False, False)
                    Next ss
                End If
            End If
        Next i
    End If
    
    ''''UDR Fuse
    If (gB_findUDR_flag) Then
        For i = 0 To UBound(UDRFuse.Category)
            m_algorithm = LCase(UDRFuse.Category(i).algorithm)
            If (m_algorithm = "base" Or m_algorithm = "vddbin") Then
                m_defreal = LCase(UDRFuse.Category(i).Default_Real)
                m_catename = UDRFuse.Category(i).Name
                m_resolution = UDRFuse.Category(i).Resoultion
                m_defval = UDRFuse.Category(i).DefaultValue
                m_bitwidth = UDRFuse.Category(i).BitWidth
                m_defval_calc = 0 ''''<MUST>
                m_updateFlag = False

                If (m_defreal = "default") Then
                'If (m_defreal Like "safe*voltage") Then
                    If (m_algorithm = "base") Then
                        m_tmpVal = m_defval / m_resolution - 1 ''''<NOTICE>
                    ElseIf (m_algorithm = "vddbin") Then
                        m_tmpVal = (m_defval - gD_BaseVoltage) / m_resolution
                    End If
                    m_delta = Abs(m_tmpVal - Int(m_tmpVal))
                    If (m_delta > 0) Then
                        m_updateFlag = False
                        TheExec.Datalog.WriteComment "<Error> " + m_catename + ":: Default Value (" & m_defval & ") is Not multiple of the resolution (" & m_resolution & ")"
                        GoTo errHandler
                    Else
                        m_updateFlag = True
                        ''''m_delta=0, here using "default" to get the binarr with the calculated value
                        m_defval_calc = auto_checkDefaultValue(m_tmpVal, m_algorithm, m_binarr, m_bitwidth, "safe voltage")

                        'm_defval_calc = auto_checkDefaultValue(m_tmpVal, m_binarr, m_bitwidth, "default")
                        UDRFuse.Category(i).DefValBitArr = m_binarr
                    End If
                Else
                    ''''other cases (bincut,default,decimal)
                    ''''do Nothing
                End If
                
                If (m_updateFlag = True) Then
                    ''''Update the correct default value to the Write structure
                    For Each ss In TheExec.sites.Existing
                        Call auto_eFuse_SetWriteDecimal("UDR", m_catename, m_defval_calc, False, False)
                    Next ss
                End If
            End If
        Next i
    End If

    ''''UDR_E Fuse, 20171103
    If (gB_findUDRE_flag) Then
        For i = 0 To UBound(UDRE_Fuse.Category)
            m_algorithm = LCase(UDRE_Fuse.Category(i).algorithm)
            If (m_algorithm = "base" Or m_algorithm = "vddbin") Then
                m_defreal = LCase(UDRE_Fuse.Category(i).Default_Real)
                m_catename = UDRE_Fuse.Category(i).Name
                m_resolution = UDRE_Fuse.Category(i).Resoultion
                m_defval = UDRE_Fuse.Category(i).DefaultValue
                m_bitwidth = UDRE_Fuse.Category(i).BitWidth
                m_defval_calc = 0 ''''<MUST>
                m_updateFlag = False

                If (m_defreal = "default") Then
                'If (m_defreal Like "safe*voltage") Then
                    If (m_algorithm = "base") Then
                        m_tmpVal = m_defval / m_resolution - 1 ''''<NOTICE>
                    ElseIf (m_algorithm = "vddbin") Then
                        m_tmpVal = (m_defval - gD_BaseVoltage) / m_resolution
                    End If
                    m_delta = Abs(m_tmpVal - Int(m_tmpVal))
                    If (m_delta > 0) Then
                        m_updateFlag = False
                        TheExec.Datalog.WriteComment "<Error> " + m_catename + ":: Default Value (" & m_defval & ") is Not multiple of the resolution (" & m_resolution & ")"
                        GoTo errHandler
                    Else
                        m_updateFlag = True
                        ''''m_delta=0, here using "default" to get the binarr with the calculated value
                        m_defval_calc = auto_checkDefaultValue(m_tmpVal, m_algorithm, m_binarr, m_bitwidth, "safe voltage")
                        'm_defval_calc = auto_checkDefaultValue(m_tmpVal, m_binarr, m_bitwidth, "default")
                        UDRE_Fuse.Category(i).DefValBitArr = m_binarr
                    End If
                Else
                    ''''other cases (bincut,default,decimal)
                    ''''do Nothing
                End If
                
                If (m_updateFlag = True) Then
                    ''''Update the correct default value to the Write structure
                    For Each ss In TheExec.sites.Existing
                        Call auto_eFuse_SetWriteDecimal("UDRE", m_catename, m_defval_calc, False, False)
                    Next ss
                End If
            End If
        Next i
    End If

    ''''UDR_P Fuse, 20171103
    If (gB_findUDRP_flag) Then
        For i = 0 To UBound(UDRP_Fuse.Category)
            m_algorithm = LCase(UDRP_Fuse.Category(i).algorithm)
            If (m_algorithm = "base" Or m_algorithm = "vddbin") Then
                m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
                m_catename = UDRP_Fuse.Category(i).Name
                m_resolution = UDRP_Fuse.Category(i).Resoultion
                m_defval = UDRP_Fuse.Category(i).DefaultValue
                m_bitwidth = UDRP_Fuse.Category(i).BitWidth
                m_defval_calc = 0 ''''<MUST>
                m_updateFlag = False

                 If (m_defreal = "default") Then
                'If (m_defreal Like "safe*voltage") Then
                    If (m_algorithm = "base") Then
                        m_tmpVal = m_defval / m_resolution - 1 ''''<NOTICE>
                    ElseIf (m_algorithm = "vddbin") Then
                        m_tmpVal = (m_defval - gD_BaseVoltage) / m_resolution
                    End If
                    m_delta = Abs(m_tmpVal - Int(m_tmpVal))
                    If (m_delta > 0) Then
                        m_updateFlag = False
                        TheExec.Datalog.WriteComment "<Error> " + m_catename + ":: Default Value (" & m_defval & ") is Not multiple of the resolution (" & m_resolution & ")"
                        GoTo errHandler
                    Else
                        m_updateFlag = True
                        ''''m_delta=0, here using "default" to get the binarr with the calculated value
                        m_defval_calc = auto_checkDefaultValue(m_tmpVal, m_algorithm, m_binarr, m_bitwidth, "safe voltage")

                        'm_defval_calc = auto_checkDefaultValue(m_tmpVal, m_binarr, m_bitwidth, "default")
                        UDRP_Fuse.Category(i).DefValBitArr = m_binarr
                    End If
                Else
                    ''''other cases (bincut,default,decimal)
                    ''''do Nothing
                End If
                
                If (m_updateFlag = True) Then
                    ''''Update the correct default value to the Write structure
                    For Each ss In TheExec.sites.Existing
                        Call auto_eFuse_SetWriteDecimal("UDRP", m_catename, m_defval_calc, False, False)
                    Next ss
                End If
            End If
        Next i
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_CMPConstant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CMPConstant_Initialize"

    Dim i As Long
    Dim ss As Variant
    Dim m_algorithm As String
    Dim m_catename As String
    Dim m_resolution As Double
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_len As Long
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim m_defvalhex As String ''''without prefix "0x"
    
    ''''Get max length of category name
    gI_CMP_catename_maxLen = 0
    For i = 0 To UBound(CMPFuse.Category)
        m_len = Len(CMPFuse.Category(i).Name)
        m_MSBBit = CMPFuse.Category(i).MSBbit

        If (m_len > gI_CMP_catename_maxLen) Then
            gI_CMP_catename_maxLen = m_len
        End If
    Next i
    gI_CMP_catename_maxLen = gI_CMP_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------
    
    ''''20160804 <User Maintain> depends on the pattern's comment
    gS_CMP_PatBitOrder = "LSB"
    gL_CMP_DigCapBits_Num = CMPFuse.Category(UBound(CMPFuse.Category)).MSBbit + 1

    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment funcName + "::"
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment FormatNumeric("   CMP_PatBitOrder = ", 25) + FormatNumeric(gS_CMP_PatBitOrder, -10)
    TheExec.Datalog.WriteComment FormatNumeric("CMP_DigCapBits_Num = ", 25) + FormatNumeric(gL_CMP_DigCapBits_Num, -10)
    TheExec.Datalog.WriteComment "---------------------------------------------"
    ''TheExec.Datalog.WriteComment ""

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20171103 add
Public Function auto_UDRE_Constant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRE_Constant_Initialize"

    Dim i As Long
    Dim ss As Variant
    Dim m_algorithm As String
    Dim m_catename As String
    Dim m_resolution As Double
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_usiLSBcycle As Long
    Dim m_usiMSBcycle As Long
    Dim m_usoLSBcycle As Long
    Dim m_usoMSBcycle As Long
    Dim m_len As Long
    Dim m_defval As Variant
    Dim m_usipretmp As String
    Dim m_usopretmp As String
    Dim m_defreal As String
    Dim m_bitwidth As Long
    Dim m_defvalhex As String ''''without prefix "0x"
    Dim m_binarr() As Long
    Dim ms_defval As New SiteVariant

    ''''Get max length of category name
    gI_UDRE_catename_maxLen = 0
    For i = 0 To UBound(UDRE_Fuse.Category)
        m_len = Len(UDRE_Fuse.Category(i).Name)
        m_MSBBit = UDRE_Fuse.Category(i).MSBbit

        If (m_len > gI_UDRE_catename_maxLen) Then
            gI_UDRE_catename_maxLen = m_len
        End If
    Next i
    gI_UDRE_catename_maxLen = gI_UDRE_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------

    m_usipretmp = ""
    m_usopretmp = ""

    For i = 0 To UBound(UDRE_Fuse.Category)
        m_catename = UDRE_Fuse.Category(i).Name
        m_algorithm = LCase(UDRE_Fuse.Category(i).algorithm)
        m_defval = UDRE_Fuse.Category(i).DefaultValue
        m_resolution = UDRE_Fuse.Category(i).Resoultion
        m_LSBbit = UDRE_Fuse.Category(i).LSBbit
        m_MSBBit = UDRE_Fuse.Category(i).MSBbit
        m_usiLSBcycle = UDRE_Fuse.Category(i).USILSBCycle
        m_usiMSBcycle = UDRE_Fuse.Category(i).USIMSBCycle
        m_usoLSBcycle = UDRE_Fuse.Category(i).USOLSBCycle
        m_usoMSBcycle = UDRE_Fuse.Category(i).USOMSBCycle
        m_defreal = LCase(UDRE_Fuse.Category(i).Default_Real)
        m_bitwidth = UDRE_Fuse.Category(i).BitWidth

        ''''-------------------------------------------------------
        If (m_usiLSBcycle < m_usiMSBcycle) Then
            gL_UDRE_USI_PatBitOrder = "LSB"
        ElseIf (m_usiLSBcycle > m_usiMSBcycle) Then
            gL_UDRE_USI_PatBitOrder = "MSB"
        End If
        If (m_usipretmp <> "" And gL_UDRE_USI_PatBitOrder <> m_usipretmp) Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: " + m_catename + " the UDRE_USI bit order is different from others"
            GoTo errHandler
        End If
        m_usipretmp = gL_UDRE_USI_PatBitOrder
        ''''-------------------------------------------------------
        
        ''''-------------------------------------------------------
        If (m_usoLSBcycle < m_usoMSBcycle) Then
            gL_UDRE_USO_PatBitOrder = "LSB"
        ElseIf (m_usoLSBcycle > m_usoMSBcycle) Then
            gL_UDRE_USO_PatBitOrder = "MSB"
        End If
        If (m_usopretmp <> "" And gL_UDRE_USO_PatBitOrder <> m_usopretmp) Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: " + m_catename + " the UDRE_USO bit order is different from others"
            GoTo errHandler
        End If
        m_usopretmp = gL_UDRE_USO_PatBitOrder
        ''''-------------------------------------------------------

        ''''20171211, Here it's used to judge default value if is Hex or Binary
        m_defval = auto_checkDefaultValue(m_defval, m_algorithm, m_binarr, m_bitwidth, m_defreal)

        'm_defval = auto_checkDefaultValue(m_defval, m_binarr, m_bitwidth, m_defreal)
        UDRE_Fuse.Category(i).DefValBitArr = m_binarr
        ms_defval = m_defval

        If (m_algorithm = "base") Then
            If (m_defreal = "decimal") Then ''''20160624 update
                gD_UDRE_VBaseFuse = m_defval
                If (m_resolution = 0) Then
                    gD_UDRE_BaseStepVoltage = 25#
                Else
                    gD_UDRE_BaseStepVoltage = m_resolution
                End If
                gD_UDRE_BaseVoltage = (gD_UDRE_VBaseFuse + 1) * gD_UDRE_BaseStepVoltage
            Else
                gD_UDRE_BaseVoltage = m_defval
                gD_UDRE_BaseStepVoltage = m_resolution
                gD_UDRE_VBaseFuse = (gD_UDRE_BaseVoltage / gD_UDRE_BaseStepVoltage) - 1
            End If

        ElseIf (m_algorithm = "fuse") Then
            gL_UDRE_eFuse_Revision = m_defval
            For Each ss In TheExec.sites.Existing
                Call auto_eFuse_SetWriteDecimal("UDRE", m_catename, m_defval, False)
            Next ss
        
        ''ElseIf (m_algorithm = "lastreserve") Then
        ElseIf (i = UBound(UDRE_Fuse.Category)) Then
            'gL_UDRE_USI_DigSrcBits_Num
            If (m_LSBbit <= m_MSBBit) Then
                gL_UDRE_USI_DigSrcBits_Num = m_MSBBit + 1
                gL_UDRE_USO_DigCapBits_Num = m_MSBBit + 1
            Else
                gL_UDRE_USI_DigSrcBits_Num = m_LSBbit + 1
                gL_UDRE_USO_DigCapBits_Num = m_LSBbit + 1
            End If
        End If
    
        ''''20160328 update, 20160624 update
        If (m_defreal <> "real" And m_defreal <> "bincut") Then
            UDRE_Fuse.Category(i).DefaultValue = m_defval
            UDRE_Fuse.Category(i).HiLMT = m_defval
            UDRE_Fuse.Category(i).LoLMT = m_defval
            If (m_algorithm = "base") Then
                If (m_defreal = "default") Then
                'If (m_defreal Like "safe*voltage") Then
                    m_defval = gD_UDRE_VBaseFuse
                    ms_defval = gD_UDRE_VBaseFuse
                End If
            End If
            ''''20180712, will do the update later on for "vddbin: safe voltage"
            ''''''''''''' please see auto_precheck_SafeVoltage_Base_VddBin()
        End If
        'If (m_algorithm = "vddbin" And m_defreal <> "decimal") Then
        'If (m_algorithm = "vddbin" And m_defreal Like "safe*voltage") Then
        If (m_algorithm = "vddbin" And m_defreal = "default") Then
        Else
            Call auto_eFuse_SetWriteVariable_SiteAware("UDRE", m_catename, ms_defval, False)
        End If
    Next i

    'Set eFuse Global Data initial
    For Each ss In TheExec.sites.Existing
        gS_UDRE_USI_BitStr(ss) = ""
    Next ss
    
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment funcName + "::"
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment FormatNumeric("UDRE_USI_PatBitOrder = ", 35) + FormatNumeric(gL_UDRE_USI_PatBitOrder, -10)
    TheExec.Datalog.WriteComment FormatNumeric("UDRE_USI_DigSrcBits_Num = ", 35) + FormatNumeric(gL_UDRE_USI_DigSrcBits_Num, -10)
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment FormatNumeric("UDRE_USO_PatBitOrder = ", 35) + FormatNumeric(gL_UDRE_USO_PatBitOrder, -10)
    TheExec.Datalog.WriteComment FormatNumeric("UDRE_USO_DigCapBits_Num = ", 25) + FormatNumeric(gL_UDRE_USO_DigCapBits_Num, -10)
    TheExec.Datalog.WriteComment "---------------------------------------------"
    ''TheExec.Datalog.WriteComment ""

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20171103 add
Public Function auto_UDRP_Constant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRP_Constant_Initialize"

    Dim i As Long
    Dim ss As Variant
    Dim m_algorithm As String
    Dim m_catename As String
    Dim m_resolution As Double
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_usiLSBcycle As Long
    Dim m_usiMSBcycle As Long
    Dim m_usoLSBcycle As Long
    Dim m_usoMSBcycle As Long
    Dim m_len As Long
    Dim m_defval As Variant
    Dim m_usipretmp As String
    Dim m_usopretmp As String
    Dim m_defreal As String
    Dim m_bitwidth As Long
    Dim m_defvalhex As String ''''without prefix "0x"
    Dim m_binarr() As Long
    Dim ms_defval As New SiteVariant

    ''''Get max length of category name
    gI_UDRP_catename_maxLen = 0
    For i = 0 To UBound(UDRP_Fuse.Category)
        m_len = Len(UDRP_Fuse.Category(i).Name)
        m_MSBBit = UDRP_Fuse.Category(i).MSBbit

        If (m_len > gI_UDRP_catename_maxLen) Then
            gI_UDRP_catename_maxLen = m_len
        End If
    Next i
    gI_UDRP_catename_maxLen = gI_UDRP_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------

    m_usipretmp = ""
    m_usopretmp = ""

    For i = 0 To UBound(UDRP_Fuse.Category)
        m_catename = UDRP_Fuse.Category(i).Name
        m_algorithm = LCase(UDRP_Fuse.Category(i).algorithm)
        m_defval = UDRP_Fuse.Category(i).DefaultValue
        m_resolution = UDRP_Fuse.Category(i).Resoultion
        m_LSBbit = UDRP_Fuse.Category(i).LSBbit
        m_MSBBit = UDRP_Fuse.Category(i).MSBbit
        m_usiLSBcycle = UDRP_Fuse.Category(i).USILSBCycle
        m_usiMSBcycle = UDRP_Fuse.Category(i).USIMSBCycle
        m_usoLSBcycle = UDRP_Fuse.Category(i).USOLSBCycle
        m_usoMSBcycle = UDRP_Fuse.Category(i).USOMSBCycle
        m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
        m_bitwidth = UDRP_Fuse.Category(i).BitWidth

        ''''-------------------------------------------------------
        If (m_usiLSBcycle < m_usiMSBcycle) Then
            gL_UDRP_USI_PatBitOrder = "LSB"
        ElseIf (m_usiLSBcycle > m_usiMSBcycle) Then
            gL_UDRP_USI_PatBitOrder = "MSB"
        End If
        If (m_usipretmp <> "" And gL_UDRP_USI_PatBitOrder <> m_usipretmp) Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: " + m_catename + " the UDRP_USI bit order is different from others"
            GoTo errHandler
        End If
        m_usipretmp = gL_UDRP_USI_PatBitOrder
        ''''-------------------------------------------------------
        
        ''''-------------------------------------------------------
        If (m_usoLSBcycle < m_usoMSBcycle) Then
            gL_UDRP_USO_PatBitOrder = "LSB"
        ElseIf (m_usoLSBcycle > m_usoMSBcycle) Then
            gL_UDRP_USO_PatBitOrder = "MSB"
        End If
        If (m_usopretmp <> "" And gL_UDRP_USO_PatBitOrder <> m_usopretmp) Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: " + m_catename + " the UDRP_USO bit order is different from others"
            GoTo errHandler
        End If
        m_usopretmp = gL_UDRP_USO_PatBitOrder
        ''''-------------------------------------------------------

        ''''20171211, Here it's used to judge default value if is Hex or Binary
        m_defval = auto_checkDefaultValue(m_defval, m_algorithm, m_binarr, m_bitwidth, m_defreal)
        'm_defval = auto_checkDefaultValue(m_defval, m_binarr, m_bitwidth, m_defreal)
        UDRP_Fuse.Category(i).DefValBitArr = m_binarr
        ms_defval = m_defval

        If (m_algorithm = "base") Then
            If (m_defreal = "decimal") Then ''''20160624 update
                gD_UDRP_VBaseFuse = m_defval
                If (m_resolution = 0) Then
                    gD_UDRP_BaseStepVoltage = 25#
                Else
                    gD_UDRP_BaseStepVoltage = m_resolution
                End If
                gD_UDRP_BaseVoltage = (gD_UDRP_VBaseFuse + 1) * gD_UDRP_BaseStepVoltage
            Else
                gD_UDRP_BaseVoltage = m_defval
                gD_UDRP_BaseStepVoltage = m_resolution
                gD_UDRP_VBaseFuse = (gD_UDRP_BaseVoltage / gD_UDRP_BaseStepVoltage) - 1
            End If

        ElseIf (m_algorithm = "fuse") Then
            gL_UDRP_eFuse_Revision = m_defval
            For Each ss In TheExec.sites
                Call auto_eFuse_SetWriteDecimal("UDRP", m_catename, m_defval, False)
            Next ss
        
        ''ElseIf (m_algorithm = "lastreserve") Then
        ElseIf (i = UBound(UDRP_Fuse.Category)) Then
            'gL_UDRP_USI_DigSrcBits_Num
            If (m_LSBbit <= m_MSBBit) Then
                gL_UDRP_USI_DigSrcBits_Num = m_MSBBit + 1
                gL_UDRP_USO_DigCapBits_Num = m_MSBBit + 1
            Else
                gL_UDRP_USI_DigSrcBits_Num = m_LSBbit + 1
                gL_UDRP_USO_DigCapBits_Num = m_LSBbit + 1
            End If
        End If
    
        ''''20160328 update, 20160624 update
        If (m_defreal <> "real" And m_defreal <> "bincut") Then
            UDRP_Fuse.Category(i).DefaultValue = m_defval
            UDRP_Fuse.Category(i).HiLMT = m_defval
            UDRP_Fuse.Category(i).LoLMT = m_defval
            If (m_algorithm = "base") Then
                If (m_defreal = "default") Then
                'If (m_defreal Like "safe*voltage") Then
                    m_defval = gD_UDRP_VBaseFuse
                    ms_defval = gD_UDRP_VBaseFuse
                End If
            End If
            ''''20180712, will do the update later on for "vddbin: safe voltage"
            ''''''''''''' please see auto_precheck_SafeVoltage_Base_VddBin()
        End If
        'If (m_algorithm = "vddbin" And m_defreal <> "decimal") Then
        'If (m_algorithm = "vddbin" And m_defreal Like "safe*voltage") Then
         If (m_algorithm = "vddbin" And m_defreal = "default") Then
        Else
            'For Each ss In TheExec.Sites.Existing
                Call auto_eFuse_SetWriteVariable_SiteAware("UDRP", m_catename, ms_defval, False)
            'Next ss
        End If
    Next i

    'Set eFuse Global Data initial
    For Each ss In TheExec.sites.Existing
        gS_UDRP_USI_BitStr(ss) = ""
    Next ss
    
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment funcName + "::"
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment FormatNumeric("UDRP_USI_PatBitOrder = ", 35) + FormatNumeric(gL_UDRP_USI_PatBitOrder, -10)
    TheExec.Datalog.WriteComment FormatNumeric("UDRP_USI_DigSrcBits_Num = ", 35) + FormatNumeric(gL_UDRP_USI_DigSrcBits_Num, -10)
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment FormatNumeric("UDRP_USO_PatBitOrder = ", 35) + FormatNumeric(gL_UDRP_USO_PatBitOrder, -10)
    TheExec.Datalog.WriteComment FormatNumeric("UDRP_USO_DigCapBits_Num = ", 35) + FormatNumeric(gL_UDRP_USO_DigCapBits_Num, -10)
    TheExec.Datalog.WriteComment "---------------------------------------------"
    ''TheExec.Datalog.WriteComment ""

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20171103 add
Public Function auto_CMPE_Constant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CMPE_Constant_Initialize"

    Dim i As Long
    Dim ss As Variant
    Dim m_algorithm As String
    Dim m_catename As String
    Dim m_resolution As Double
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_len As Long
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim m_defvalhex As String ''''without prefix "0x"
    
    ''''Get max length of category name
    gI_CMPE_catename_maxLen = 0
    For i = 0 To UBound(CMPE_Fuse.Category)
        m_len = Len(CMPE_Fuse.Category(i).Name)
        m_MSBBit = CMPE_Fuse.Category(i).MSBbit

        If (m_len > gI_CMPE_catename_maxLen) Then
            gI_CMPE_catename_maxLen = m_len
        End If
    Next i
    gI_CMPE_catename_maxLen = gI_CMPE_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------
    
    ''''20160804 <User Maintain> depends on the pattern's comment
    gS_CMPE_PatBitOrder = "LSB"
    gL_CMPE_DigCapBits_Num = CMPE_Fuse.Category(UBound(CMPE_Fuse.Category)).MSBbit + 1

    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment funcName + "::"
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment FormatNumeric("CMPE    CMPE_PatBitOrder = ", 35) + FormatNumeric(gS_CMPE_PatBitOrder, -10)
    TheExec.Datalog.WriteComment FormatNumeric("CMPE CMPE_DigCapBits_Num = ", 35) + FormatNumeric(gL_CMPE_DigCapBits_Num, -10)
    TheExec.Datalog.WriteComment "---------------------------------------------"
    ''TheExec.Datalog.WriteComment ""

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20171103 add
Public Function auto_CMPP_Constant_Initialize()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CMPP_Constant_Initialize"

    Dim i As Long
    Dim ss As Variant
    Dim m_algorithm As String
    Dim m_catename As String
    Dim m_resolution As Double
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_len As Long
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim m_defvalhex As String ''''without prefix "0x"
    
    ''''Get max length of category name
    gI_CMPP_catename_maxLen = 0
    For i = 0 To UBound(CMPP_Fuse.Category)
        m_len = Len(CMPP_Fuse.Category(i).Name)
        m_MSBBit = CMPP_Fuse.Category(i).MSBbit

        If (m_len > gI_CMPP_catename_maxLen) Then
            gI_CMPP_catename_maxLen = m_len
        End If
    Next i
    gI_CMPP_catename_maxLen = gI_CMPP_catename_maxLen + 2 ''''with additional 2 spaces
    ''''-------------------------------------------------------
    
    ''''20160804 <User Maintain> depends on the pattern's comment
    gS_CMPP_PatBitOrder = "LSB"
    gL_CMPP_DigCapBits_Num = CMPP_Fuse.Category(UBound(CMPP_Fuse.Category)).MSBbit + 1

    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment funcName + "::"
    TheExec.Datalog.WriteComment "---------------------------------------------"
    TheExec.Datalog.WriteComment FormatNumeric("CMPP    CMPP_PatBitOrder = ", 35) + FormatNumeric(gS_CMPP_PatBitOrder, -10)
    TheExec.Datalog.WriteComment FormatNumeric("CMPP CMPP_DigCapBits_Num = ", 35) + FormatNumeric(gL_CMPP_DigCapBits_Num, -10)
    TheExec.Datalog.WriteComment "---------------------------------------------"
    ''TheExec.Datalog.WriteComment ""

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

