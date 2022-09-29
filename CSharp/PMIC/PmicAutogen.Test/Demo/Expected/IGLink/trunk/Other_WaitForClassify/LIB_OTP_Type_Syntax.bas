Attribute VB_Name = "LIB_OTP_Type_Syntax"
'T-AutoGen-Version : 1.3.0.1
'ProjectName_A1_TestPlan_20220226.xlsx
'ProjectName_A0_otp_AVA.otp
'ProjectName_A0_OTP_register_map.yaml
'ProjectName_A0_Pattern_List_Ext_20190823.csv
'ProjectName_A0_scgh_file#1_20200207.xlsx
'ProjectName_A0_VBTPOP_Gen_tool_MP10P_BuckSW_UVI80_DiffMeter_20200430.xlsm

Option Explicit
'**************************************************************************************************
'Owners need to maintain all the contents in this VB module
'**************************************************************************************************

'**************************************************************************************************
'***Autogen Note: 20181004***
'**************************************************************************************************
''   Total 4 sheets output from Autogen:
''   1. OTP_register_map: generated by inputed "xxx.yaml" and "xxx.otp"
''   2. AHB_register_map: copy "chip_APC_MPXP_register_map#xx.xlsx" directly.
''   3. otp_default_reglist: convert the decimal default values in "OTP_register_map" to Hex. '1248 OTP_reg_categories on MP4T
''   4. GlobalAddressMap: Keep the unique "reg_name" and its "reg_address" from "AHB_register_map" '2595 unique reg_name on MP4T
'***************************************************************************************************

'==========================================================================
'OWNER:FUJI/TER/ORANGE
'==========================================================================
    '(A):
    
'OWNER: Orange
'(Need to update based on the otp_reg_map assignment)
      

            
    '___Device Burn Status
    Public Const g_sOTP_SHEETNAME As String = "OTP_register_map"   'OTP register Table Sheet.

    Public Const g_iOTP_ADDR_START As Integer = 0                   'ADDR.[4095:3584].
    Public Const g_iOTP_ADDR_END  As Integer = 1023                 'ADDR.[4095:3584].
    Public Const g_iOTP_ADDR_TOTAL  As Integer = 1024               'g_iOTP_ADDR_TOTAL:ADDR. number=512'''ADDR.[511:0].==>'ADDR.[4095:3584].
    Public Const g_iOTP_ADDR_OFFSET  As Integer = &HC00             '&H000 &HE00
    
    '___if OTP_BLOCK_B0 in "OTP_register_map" is always 0, no neccessay to modify "g_iOTP_BLOCK_B0_NUM" and "g_iOTP_BITS_PERBLOCK"
    Public Const g_iOTP_BLOCK_B0_NUM  As Integer = 1 'check "otp_b0" in OTP_register_map, if always 0, then g_iOTP_BLOCK_B0_NUM=1.
    Public Const g_iOTP_BITS_PERBLOCK  As Integer = 32 '8 in MP1P
    
    Public Const g_iOTP_ADDR_BW  As Integer = 16
    Public Const g_iOTP_DATA_BW = g_iOTP_BLOCK_B0_NUM * g_iOTP_BITS_PERBLOCK '3 B0_num and 8 bitsPerBlock in MP3P
    
    Public Const g_iOTP_MACRONUM  As Integer = 4
    Public Const g_iOTP_REGDATA_BW As Integer = g_iOTP_BITS_PERBLOCK
    
    Public Const g_sOTPDATA_FILEDIR As String = ".\OTPDATA\"
    Public Const g_iAHB_CRCSIZE  As Integer = 4441 ' 1853 for MP3P, updated as 1247 for MP4T on 20180925
    Public g_iOvwCrcCnt As Integer '= 131 '42
    '___VPP & VDD
    Public Const g_sVPP_PINNAME As String = "VPP_DC30"
    Public Const g_sVDD_PINNAME As String = "VDD_MAIN_UVI80"
    Public Const g_bVPP_DISABLE As Boolean = False 'Set false to execute OTP burn
    
    '___OTP_REGISTER_NAME:
    Public Const g_sOTP_PRGM_BIT_REG As String = "OTP__ARRAY_PROGRAMMED_0"  'Lock bit check
    Public Const g_sOTP_ECID_BIT_REG As String = "OTP_HOST_INTERFACE_CHIP_ID_0_3"
    Public Const g_sOTP_CRC_BIT_REG As String = "OTP__CRC_0"
'    Public Const gS_OTP_HOST_INTERFACE_I2C As String = "OTP_HOST_INTERFACE_I2C_ADDR_11"
'    Public Const gS_OTP_REV = "OTP_HOST_INTERFACE_OTP_REVISION_1"
    
    '___Define OTP Chip Name array in InitializeOtpTable
    Public g_asChipIDName() As Variant
    
    '___Define OTP Chip Name array in InitializeOtpTable
    Public g_asOTPRevName() As Variant
    
    '___OTP version and SVN version register definitions
    Public Const g_sOTP_TPVERSION_M As String = "OTP_OTP_SLV_MAJOR_OTP_VERSION_4432" '"OTP_OTP_SLV_MAJOR_OTP_VERSION_1244"
    Public Const g_sOTP_TPVERSION_S As String = "OTP_OTP_SLV_MINOR_OTP_VERSION_4432" '"OTP_OTP_SLV_MINOR_OTP_VERSION_1244"
    Public Const g_sSVN_VERSION_MSB As String = "OTP_OTP_SLV_LCK0_4433" '"OTP_OTP_SLV_LCK0_1245"
    Public Const g_sSVN_VERSION_LSB As String = "OTP_OTP_SLV_LCK1_4434" '"OTP_OTP_SLV_LCK1_1246"
   
    Public Const g_sAHB_WRTIE_TEST_PAT As String = "DD_SERA0_C_IN00_JT_XXXX_PFF_JTG_UNS_ALLFV_SI_TC2AHB_WRDSC" '"DD_SUZA0_C_IN00_JT_XXXX_PFF_JTG_UNS_ALLFV_SI_TC2AHB_WRDSC"
    Public Const g_sAHBBF_WRTIE_TEST_PAT As String = "DD_SERA0_C_IN00_JT_XXXX_PFF_JTG_UNS_ALLFV_SI_TC2AHBBF_WRDSC" ' "DD_SUZA0_C_IN00_JT_XXXX_PFF_JTG_UNS_ALLFV_SI_TC2AHBBF_WRDSC"
    Public Const g_sTDI As String = "GPIO2"
    Public Const g_sTDO As String = "GPIO4"
    Public Const g_sDIG_PINS As String = "ALL_DIG_PINS_NO_FRC"
    
    '___IO Pins example in MP1P
    #If 0 Then
    Public Const g_sOTP_ADDR_PINS As String = "GPIO16,GPIO15,GPIO14,GPIO13,GPIO12,GPIO11,GPIO10,GPIO9,GPIO8,GPIO7,GPIO6"
    Public Const g_sOTP_DATA_PINS As String = "GPIO24,GPIO23,GPIO22,GPIO21,GPIO20,GPIO19,GPIO18,GPIO17"
    #End If
    
     Public Const g_sAHB_SHEETNAME As String = "AHB_register_map" 'AHB register Table Sheet.
    '___AHB Related Parameters
    Public Const g_iAHB_ADDR_BW  As Integer = 16
    Public Const g_iAHB_BW  As Integer = 8
    Public Const g_iAHB_DATA_BW As Integer = 8
    Public g_RegVal As New SiteLong
    
'    '___OTP FW PATTERNS:
'    Public g_sPatDesignReg As String '= ".\Patterns\CPU\ANALOGUE\DD_AVSA0_C_IN00_AN_XXXX_PFF_JTG_UNS_ALLFV_SI_DESIGNRG_APZFWTSU_2_A0_1901231059.PAT"
'    Public g_sPatSystemReg As String '= ".\Patterns\CPU\ANALOGUE\DD_AVSA0_C_IN00_AN_XXXX_PFF_JTG_UNS_ALLFV_SI_SYSTEMRG_APZFWTSU_2_A0_1901231059.PAT"
'    Public g_sFWOTP_PAT    As String
    
    '___OTP OneShot Patterns:
    Public Const g_sOTP_ONESHOT_WRITE As String = "OTP_WRITE_ALL_DSC"
    Public Const g_sOTP_ONESHOT_READ As String = "OTP_READ_ALL_DSC"
    
    '___AHB PATTERNS:
    Public Const g_sAHB_WRITE_PAT As String = "DD_SERA0_C_IN00_JT_XXXX_PFF_JTG_UNS_ALLFRV_SI_TC2AHB_WRDSC" '"DD_SUZA0_C_IN00_JT_XXXX_PFF_JTG_UNS_ALLFV_SI_TC2AHB_WRDSC"
    Public Const g_sAHBBF_WRITE_PAT As String = "DD_SERA0_C_IN00_JT_XXXX_PFF_JTG_UNS_ALLFRV_SI_TC2AHBBF_WRDSC" '"DD_SUZA0_C_IN00_JT_XXXX_PFF_JTG_UNS_ALLFV_SI_TC2AHBBF_WRDSC"
    Public Const g_sAHB_READ_PAT As String = "DD_SERA0_C_IN00_JT_XXXX_PFF_JTG_UNS_ALLFRV_SI_TC2AHB_RDDSC" '"DD_SUZA0_C_IN00_JT_XXXX_PFF_JTG_UNS_ALLFV_SI_TC2AHB_RDDSC"
    Public Const g_sAHB_SEL As String = "PP_SERA0_C_IN00_JT_XXXX_WIR_JTG_UNS_ALLFRV_SI_AHB_SEL"
    
    #If 1 Then
    Public Const g_sOTP_WRITE_SELF_TEST_PAT As String = "DD_SERA0_C_IN00_EF_XXXX_EFP_JTG_PRG_ALLFRV_SI_OTPDEBWR_TSU"
    Public Const g_sOTP_READ_SELF_TEST_PAT As String = "DD_SERA0_C_IN01_EF_XXXX_EFP_JTG_PRG_ALLFRV_SI_OTPDEBWR_TSU"
    #End If
    
    '___AHB vs OTP Check
    Public Const g_sOTP_OWNER_FOR_CHECK_AHB As String = "system,design,trim" 'Remove sival,ate,trim; Filter by OTP_Owner
    Public Const g_sOTP_OWNER_FOR_CHECK As String = "system,design,trim"
    Public Const g_sOTP_DEFAULT_REAL_FOR_PRECHECK As String = "Real"
    
    '2018/10/04:Comfirm By DE: Remove all except ECID
    '___Used in test_OTP_CHECK_DefaulReal
    Public Const g_asOTP_NO_NEED_TO_SET As String = "OTP_HOST_INTERFACE_CHIP_ID_0_3,OTP_HOST_INTERFACE_CHIP_ID_1_4,OTP_HOST_INTERFACE_CHIP_ID_2_5,OTP_HOST_INTERFACE_CHIP_ID_3_6,OTP_HOST_INTERFACE_CHIP_ID_4_7," _
                                              & "OTP_HOST_INTERFACE_CHIP_ID_5_8,OTP_HOST_INTERFACE_CHIP_ID_6_9,OTP_HOST_INTERFACE_CHIP_ID_7_10"
    '2018/06/26: Allow AHB vs OTP mis-match
    '___Used in Auto_AHB_OTP_WriteCheckByCondition
    Public Const g_asBYPASS_AHBOTP_CHECK As String = "OTP_HOST_INTERFACE_CHIP_ID_0_3,OTP_HOST_INTERFACE_CHIP_ID_1_4,OTP_HOST_INTERFACE_CHIP_ID_2_5,OTP_HOST_INTERFACE_CHIP_ID_3_6,OTP_HOST_INTERFACE_CHIP_ID_4_7," _
                                                & "OTP_HOST_INTERFACE_CHIP_ID_5_8,OTP_HOST_INTERFACE_CHIP_ID_6_9,OTP_HOST_INTERFACE_CHIP_ID_7_10"
    
'==========================================================================
'OWNER: Orange
'(Update by the real wafer size and setxy function here)
'==========================================================================
    Public Const g_iXCOORD_LOW_LMT As Integer = 0
    Public Const g_iXCOORD_HI_LMT As Integer = 34 + 1
    Public Const g_iYCOORD_LOW_LMT As Integer = 0
    Public Const g_iYCOORD_HI_LMT As Integer = 30 + 1
'==========================================================================
    'Lot Number[0:35]
    Public Const g_iLOTID_BITS_START As Integer = 0
    Public Const g_iLOTID_BITS_BW  As Integer = 36
    'Wafer ID[36:40]
    Public Const g_iWFID_BITS_START  As Integer = 36
    Public Const g_iWFID_BITS_BW  As Integer = 5
    'X Coordinate[41:46]
    Public Const g_iXCOORD_BITS_START  As Integer = 41
    Public Const g_iXCOORD_BITS_BW  As Integer = 6
    'Y Coordinate[47:52]
    Public Const g_iYCOORD_BITS_START  As Integer = 47
    Public Const g_iYCOORD_BITS_BW As Integer = 6
    'OTP Version[56:58]
    Public Const g_iOTPREV_BITS_START As Integer = 56
    Public Const g_iOTPREV_BITS_BW  As Integer = 3
    
    '************************************************
    '                  OTP_Enum
    '************************************************
    Public Enum g_eRegWriteRead
        eREGWRITE = 0
        eREGREAD = 1
    End Enum

    Public Enum g_eOTPBLOCK_TYPE
        eECID_OTPBURN = 1
        eCRC_OTPBURN = 2
    End Enum
    
    Public Enum g_eHOST_INTERFACE_TYPE
        ePLATFORM_ID = 1
        eOTP_CONSUMER_TYPE = 2
        eOTP_REVISION_TYPE = 3
        eTP_OTP_VERSION = 4
    End Enum
    
    Public Enum g_eAHB_OTP_COMP_TYPE
        eCHECK_ALL = 1
        eCHECK_BY_CONDITION = 2
    End Enum
    
'************************************************
'         OTP_Type_Syntax
'************************************************
    Public Type g_tOTPCategoryParamResultSyntax
        BitStrM                 As New SiteVariant ''''(MSB...LSB), using dynamic array
        Value                   As New SiteVariant
        HexStr                  As New SiteVariant
    End Type
    
    Public Type g_tOTPDataCategoryParamSyntax
        lOtpIdx                 As Long
        sOtpRegisterName        As String      '''' OTP + 'Name' + 'Instance_Name'+ 'OTP_REG_ADD'
        sName                   As String
        sInstanceName           As String
       
        lBitWidth               As Long
        lOtpOffset              As Long        ''''Offset@BLOCK:bitfield offset in 8-bit container.
        lOtpA0                  As Long        'OTP address (0~1023)
        lOtpB0                  As Long        'OTP block (0~2 on MP1P, always 0 on others projects)
       
        lDefaultValue           As Long
        sDefaultORReal          As String
        
        lOtpBitStrStart         As Long 'OTP-DSP
        lOtpBitStrEnd           As Long 'OTP-DSP
            
        '___AHB infomation for OTP-AHB Campare
        sRegisterName            As String
        sAhbAddress              As String
        sAhbMask(g_iAHB_BW - 1)  As Long
        lAhbMaskVal           As Long
        svAhbReadVal             As New SiteVariant
        svAhbReadValByMaskOfs    As New SiteVariant
        lCalDeciAhbByMaskOfs     As Long
        lOtpRegAdd               As Long        ''''Address of OTP register.
        lOtpRegOfs               As Long        ''''Bitfield offset in OTP register.
        sOTPOwner                As String
        
        '___Store the intend to write/ read back otp register values
        Write                   As g_tOTPCategoryParamResultSyntax
        Read                    As g_tOTPCategoryParamResultSyntax
        
        'DefaultValueToBinarny   As Variant
        wBitIndex As New DSPWave
    End Type
    
    Public Type g_tOTPDataCategorySyntax
        Category()              As g_tOTPDataCategoryParamSyntax ''''using dynamic array
    End Type
    Private Type OTPRevCategoryParamSyntax
        Index                   As Long
        PKGName                 As String
        DefaultValue()          As Long
        DefaultorReal()         As String
    End Type
    Public Type OTPRevCategorySyntax
        Category()              As OTPRevCategoryParamSyntax
    End Type


    
    Public g_sSubTestMode As String
    'Public SubTestCondition As String
    Public g_sSubTestCondition As String
    
    '___Device burn status check
    Public g_sbOtped                 As New SiteBoolean 'if this is true then we have already OTPed the part
    Public g_sbOtpedECID             As New SiteBoolean 'if this is true then we have already OTPed the part
    Public g_sbOtpedPGM              As New SiteBoolean 'if this is true then we have already OTPed the part
    Public g_sbOtpedCRC              As New SiteBoolean 'if this is true then we have already OTPed the part

    '___global variant for site selector
    Public Site                         As Variant
    Public iSite                        As Variant
    Public g_lOTPCateNameMaxLen         As Long
    Public g_lOTPRevision               As Long
    Public g_lTestProgVersion           As Long 'Used in InitializeOtpVersion
    Public g_sOTPType                   As String
    Public g_sOTPRevisionType           As String
    Public g_asLogTestName()            As String
    Public g_aslOTPChipReg(7)           As New SiteLong
    'Public g_sPreTmpEnWrd              As String
    
    '___DebugPrint Flags
    Public g_bOTPRevDataUpdate          As Boolean 'Check InitializeOtpVersion status
    Public g_bEnWrdOTPFTProg            As Boolean
    Public g_bOTPFW                     As Boolean
    Public g_bOTPOneShot                As Boolean
    Public g_bTTR_ALL                   As Boolean
    Public Const g_sAHB_CHECK = False
    Public g_bFWDlogCheck               As Boolean
    Public g_bOTPcmpAHB                 As Boolean

    
    ''''The below variables stand for the result from the prober read
    Public g_sLotID   As String
    Public g_lWaferID As Long
    Public g_slXCoord As New SiteLong
    Public g_slYCoord As New SiteLong
    
    ''''The below variables stand for the result from OTP ECID read
    Public g_svOTP_LotID    As New SiteVariant
    Public g_slOTP_WaferId  As New SiteLong
    Public g_slOTP_XCoord   As New SiteLong
    Public g_slOTP_YCoord   As New SiteLong
    
    Public g_slOTP_Rev As New SiteLong
    Public g_lRevIdx   As Long

    '___add for DFT1~DFT6 (special case on MP5T combo chip)
    'Public g_sDftType As String
    

'************************************************
    'Debug Print Flags
'************************************************
    Public g_bSetWriteDebugPrint         As Boolean
    Public g_bGetWriteDebugPrint         As Boolean
    Public g_bGetReadDebugPrint          As Boolean
    Public g_bOTPRevCheckDebugPrint      As Boolean
    Public g_bOTPDsscBitsDebugPrint      As Boolean
    Public g_bAHBWriteCheckDebugPrint    As Boolean
    'Public g_bBlankFlag As Boolean
    Public g_bDump2CsvDebugPrint         As Boolean '20190611
    Public g_bTestTimeProfileDebugPrint  As Boolean
    Public g_bExpectedActualDebugPrint   As Boolean
    Public g_sOTPVerPrevious                As String





Public Sub SetOtpOvwSize()
    Dim sFuncName As String: sFuncName = "SetOtpOvwSize"
    On Error GoTo ErrHandler
    'Set default SetOtpOvwSize as 0
    'If g_sOTPRevisionType = "OTP_JPD_V01" Then
        g_iOvwCrcCnt = 0
    'End If
Exit Sub
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function SetXY(x As Integer, y As Integer) As Long
    Dim sFuncName As String: sFuncName = "setXY"
    On Error GoTo ErrHandler

    If (LCase(TheExec.CurrentChanMap) Like "*cp*1_site*") Then
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)

    ElseIf (LCase(TheExec.CurrentChanMap) Like "*x8*") Then
        'Avus: ChannelMap_X8
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, x - 3)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, x - 6)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(3, x - 9)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(4, x)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(5, x - 3)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(6, x - 6)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(7, x - 9)

        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(3, y)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(4, y - 2)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 2)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(6, y - 2)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(7, y - 2)

    Else

        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(0, x)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(1, x - 3)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(2, x - 6)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(3, x)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(4, x - 3)
        Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(5, x - 6)

        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(0, y)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(1, y)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(2, y)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(3, y - 2)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(4, y - 2)
        Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(5, y - 2)

    End If

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Sub DefineChipID()
    Dim sFuncName As String: sFuncName = "DefineChipID"
    On Error GoTo ErrHandler
    '********************User Maintain**************************
    g_asChipIDName = Array("OTP_HOST_INTERFACE_CHIP_ID_0_3", _
                                "OTP_HOST_INTERFACE_CHIP_ID_1_4", _
                                "OTP_HOST_INTERFACE_CHIP_ID_2_5", _
                                "OTP_HOST_INTERFACE_CHIP_ID_3_6", _
                                "OTP_HOST_INTERFACE_CHIP_ID_4_7", _
                                "OTP_HOST_INTERFACE_CHIP_ID_5_8", _
                                "OTP_HOST_INTERFACE_CHIP_ID_6_9", _
                                "OTP_HOST_INTERFACE_CHIP_ID_7_10")
    '********************User Maintain**************************
Exit Sub
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub DefineOtpRevName()
    Dim sFuncName As String: sFuncName = "DefineOtpRevName"
    On Error GoTo ErrHandler
    '********************User Maintain**************************
        g_asOTPRevName = Array("OTP_HOST_INTERFACE_OTP_REVISION_1", _
                               "OTP_HOST_INTERFACE_OTP_CONSUMER_1", _
                               "OTP_HOST_INTERFACE_PLATFORM_ID_2")
    '********************User Maintain**************************
Exit Sub
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + sFuncName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

