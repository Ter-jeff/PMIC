Attribute VB_Name = "VBT_LIB_Nwire_WriteRead"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

Private ModBurstStatus As Boolean
Private g_SPMI_EN As Boolean
Private g_Nwire_EN   As Boolean
Private g_BitField_EN As Boolean
Private KeepAliveStatus As Boolean
Private GNGStatus As Boolean

'Public Function AHB_WRITE_JTAG(inAddress As Long, ByVal inData As Variant, Optional inFieldMask As Long = 0) As Long
'___20200313, for AHB New Method
Public Function AHB_WRITE_JTAG(Address_In As Variant, ByVal inData As Variant, Optional Field_Mask_In As Variant = 0) As Long
    
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "AHB_WRITE_JTAG"

    Dim PortName  As String
    PortName = "NWIRE_JTAG"

    Dim TestName  As String
    Dim Address_Data As New SiteLong

    If TheExec.TesterMode = testModeOffline Then Exit Function

    Dim ShiftBits As Long
    Dim num       As Integer
    Dim tmpField_Mask As Long
    Dim MaskFiled As New SiteLong
    Dim Data      As New SiteLong
    
    '_______________________________________________________________________
    '20200313, for new AHB method
    Dim inAddress As Long
    Dim inFieldMask As Long
    If TypeName(Address_In) = "String" Then
        Call GetAHB_Add_BF_Value(CStr(Address_In), inAddress, inFieldMask)
    Else
        inAddress = CLng(Address_In)
        inFieldMask = CLng(Field_Mask_In)
    End If
    '__________________________________________________________________________
    
    ShiftBits = 0
    tmpField_Mask = inFieldMask
    For num = 1 To g_iAHB_BW
        If (tmpField_Mask And &H1) = 0 Then
            Exit For
        End If
        tmpField_Mask = Fix(tmpField_Mask / 2)
        ShiftBits = ShiftBits + 1
    Next num

    'Data = inData.ShiftLeft(ShiftBits)
    Data = inData
    Data = Data.ShiftLeft(ShiftBits)
    MaskFiled = tmpField_Mask

    TestName = TheExec.DataManager.InstanceName
    TestName = TestName & CStr(Hex(inAddress))

    With TheHdw.Protocol.ports(PortName).NWire.Frames("Write")
        .Fields("Mask").Value = inFieldMask    'MaskFiled
        .Fields("Addr").Value = inAddress
        .Fields("Data").Value = Data
        .Execute tlNWireExecutionType_Default
    End With

    'TheHdw.Protocol.Ports(PortName).IdleWait

    Exit Function

ErrHandler:
    'Debug.Print err.Description
    '   Stop  '//2019_1213
    'Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Public Function AHB_WRITE_JTAG_noBitField(inAddress As Long, ByVal inData As Variant, Optional inFieldMask As Long = 0, Optional ForceWholeWrite As Boolean = False) As Long
'___20200313, for AHB New Method
Public Function AHB_WRITE_JTAG_noBitField(Address_In As Variant, ByVal inData As Variant, Optional Field_Mask_In As Variant = 0, Optional ForceWholeWrite As Boolean = False) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "AHB_WRITE_JTAG_noBitField"

    Dim PortName  As String
    PortName = "NWIRE_JTAG"

    Dim cap_SDOVAL As SiteLong

    Dim myPLD     As New PinListData
    Dim DatafromPA As New SiteLong
    Dim PariyfromPA As New SiteLong
    Dim CMEMtransindexes As INWireCMEMTransactionIndexes
    Dim TransIndex As INWireCMEMTransactionIndex

    Dim nSite     As Variant
    Dim Address_Data As New SiteLong


    If TheExec.TesterMode = testModeOffline Then Exit Function

    Dim ShiftBits As Long
    Dim num       As Integer
    Dim tmpField_Mask As Long
    Dim MaskFiled As New SiteLong
    Dim Data      As New SiteLong
    
    '_______________________________________________________________________
    '20200313, for new AHB method
    Dim inAddress As Long
    Dim inFieldMask As Long
    If TypeName(Address_In) = "String" Then
        Call GetAHB_Add_BF_Value(CStr(Address_In), inAddress, inFieldMask)
    Else
        inAddress = CLng(Address_In)
        inFieldMask = CLng(Field_Mask_In)
    End If
    '__________________________________________________________________________
    
    Data = inData

    If inFieldMask = 0 Then
        ShiftBits = 0
        Data = Data.ShiftLeft(ShiftBits)
        'Data = inData.ShiftLeft(ShiftBits)
    Else
        Dim outData As New SiteLong
        Dim RealConcern_InData As New SiteLong
        Dim Except_ConcernData_From_ReadData As New SiteLong
        ''''-Read back to combine------------------------------------------------------------------------
        If ForceWholeWrite = False Then
            '' Special setting with Bitfield
            AHB_READ_JTAG inAddress, outData, , False
            Except_ConcernData_From_ReadData = outData.BitWiseAnd(inFieldMask)

            ShiftBits = 0
            tmpField_Mask = inFieldMask
            For num = 1 To g_iAHB_BW
                If (tmpField_Mask And &H1) = 0 Then
                    Exit For
                End If
                tmpField_Mask = Fix(tmpField_Mask / 2)
                ShiftBits = ShiftBits + 1
            Next num
            'RealConcern_InData = inData.ShiftLeft(ShiftBits).BitwiseAnd((2 ^ 8 - 1) - inFieldMask)
            RealConcern_InData = Data.ShiftLeft(ShiftBits).BitWiseAnd((2 ^ 8 - 1) - inFieldMask)
            Data = Except_ConcernData_From_ReadData.Add(RealConcern_InData)
        Else
            '' Special setting with Bitfield
            AHB_READ_JTAG inAddress, outData, , False
            Except_ConcernData_From_ReadData = outData.BitWiseAnd(inFieldMask)
            ShiftBits = 0
            'RealConcern_InData = inData.ShiftLeft(ShiftBits).BitwiseAnd((2 ^ 8 - 1) - inFieldMask)
            RealConcern_InData = Data.ShiftLeft(ShiftBits).BitWiseAnd((2 ^ 8 - 1) - inFieldMask)
            Data = Except_ConcernData_From_ReadData.Add(RealConcern_InData)
        End If
    End If


    With TheHdw.Protocol.ports(PortName).NWire.Frames("Write")
        '''            .Fields("Mask").Value = inFieldMask 'MaskFiled
        .Fields("Addr").Value = inAddress
        .Fields("Data").Value = Data
        .Execute tlNWireExecutionType_Default
    End With

    'TheHdw.Protocol.Ports(PortName).IdleWait

    Exit Function

ErrHandler:
    'Debug.Print err.Description
    '   Stop   '\\2019_1213
    'Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

' ****************************************** Example : how to use AHB_READDSC ************************************************
'
' Previous Project (Imola or SStone): AHB_READDSC BUCK0_HP2_CFG_0 , regval
'
' Suzuka Project : 1. Read Register(Same as previous Project): AHB_READDSC BUCK0_HP2_CFG_0.Addr , regval
'                2. Read Register By Field : AHB_READDSC BUCK0_HP2_CFG_0.Addr , regval, BUCK0_HP2_CFG_0.CFG_CC_BLK_SET
'
'*****************************************************************************************************************************
'2018/05/26
'Public Function AHB_READ_JTAG(inAddress As Long, outData As SiteLong, Optional Field_Mask As Long = 0, Optional bDBGlog As Boolean = True, Optional dWaitTime As Double = 10 * ms) As SiteLong
'___20200313, for AHB New Method
Public Function AHB_READ_JTAG(Address_In As Variant, outData As SiteLong, Optional Field_Mask_In As Variant = 0, Optional bDBGlog As Boolean = True, Optional dWaitTime As Double = 10 * ms) As SiteLong

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "AHB_READ_JTAG"


    Dim PortName  As String
    PortName = "NWIRE_JTAG"

    Dim myPLD_nWire As New PinListData
    'Dim indexes1 As INWireHRAMTransactionIndexes
    Dim indexes   As INWireCMEMTransactionIndexes
    Dim TempData  As New SiteDouble

    If TheExec.TesterMode = testModeOffline Then Exit Function    '20181017

    Dim ReferenceTime As Double, ElapsedTime As Double
    
    '_______________________________________________________________________
    '20200313, for new AHB method
    Dim inAddress As Long
    Dim Field_Mask As Long
    If TypeName(Address_In) = "String" Then
        Call GetAHB_Add_BF_Value(CStr(Address_In), inAddress, Field_Mask)
    Else
        inAddress = CLng(Address_In)
        Field_Mask = CLng(Field_Mask_In)
    End If
    '__________________________________________________________________________

    '    ReferenceTime = TheExec.Timer
    With TheHdw.Protocol.ports(PortName).NWire.Frames("Read")
        .Fields("Addr").Value = inAddress
        .Execute tlNWireExecutionType_CaptureInCMEM
    End With
    '    ElapsedTime = TheExec.Timer(ReferenceTime)
    '    Debug.Print ElapsedTime

    'thehdw.Protocol.Ports(PortName).IdleWait


    For Each g_Site In TheExec.Sites
        Set indexes = TheHdw.Protocol.ports(PortName).NWire.CMEM.Transactions.Read(0).Pins(PortName).Value(g_Site)
        outData = indexes(0).Fields("Data").Value
    Next g_Site

    '******************************************************************************************
    Dim BitAND As Long, Offset As Long
    Dim BitANDStr As String, CalcData As New SiteLong

    If Field_Mask > 0 Then
        BitAND = (&HFF) Xor Field_Mask
        BitANDStr = ConvertFormat_Dec2Bin_Complement(BitAND, 8)
        Offset = InStr(StrReverse(BitANDStr), "1") - 1
        CalcData = outData.BitWiseAnd(BitAND).ShiftRight(Offset)
        outData = CalcData
    End If

    For Each Site In TheExec.Sites
        'If bDBGlog = True Then TheExec.Datalog.WriteComment "Address-h'" & Hex(inAddress) & "(d'" & (inAddress And &HFF) & ")/" & "Data-" & Hex(outData(Site))
        If bDBGlog = True Then Debug.Print "Address-h'" & Hex(inAddress) & "(d'" & (inAddress And &HFF) & ")/" & "Data-" & Hex(outData(Site))
    Next Site

    Exit Function

ErrHandler:
    'Debug.Print err.Description
    ''        Stop   '//2019_1213
    'Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

'Public Function AHB_WRITE(inAddress As Long, ByVal inData As SiteLong, Optional inFieldMask As Long = 0, Optional ForceWholeWrite As Boolean = False) As Long
'___20200313, New AHB Method
Public Function AHB_WRITE(inAddress As Variant, ByVal inData As SiteLong, Optional inFieldMask As Variant = 0, Optional ForceWholeWrite As Boolean = False) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "AHB_WRITE"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    
    If g_Nwire_EN Then
        If g_SPMI_EN Then
            If g_BitField_EN Then
                SPMIBF_PA_WRITE inAddress, inData, inFieldMask
            Else
                SPMI_PA_WRITE inAddress, inData
            End If
        Else
            If g_BitField_EN Then
                AHB_WRITE_JTAG inAddress, inData, inFieldMask  ' BitField
            Else
                AHB_WRITE_JTAG_noBitField inAddress, inData, inFieldMask, ForceWholeWrite 'No BitField
            End If
        End If
    Else
        If g_BitField_EN Then
            AHB_WRITEDSC inAddress, inData, inFieldMask  ' BitField
        Else
            AHB_WRITEDSC_NoBitField inAddress, inData, inFieldMask, ForceWholeWrite 'No BitField
        End If
    End If
    
    Exit Function
    
ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Public Function AHB_READ(inAddress As Long, outData As SiteLong, Optional Field_Mask As Long = 0, Optional bDBGlog As Boolean = True, Optional dWaitTime As Double = 10 * ms) As Long
'___20200313, New AHB Method
Public Function AHB_READ(inAddress As Variant, outData As SiteLong, Optional inFieldMask As Variant = 0, Optional bDBGlog As Boolean = False, Optional dWaitTime As Double = 10 * ms) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "AHB_READ"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function
    
    If g_Nwire_EN Then
        If g_SPMI_EN Then
            SPMI_PA_READ_TO_CRAM inAddress, outData, inFieldMask, bDBGlog ', dWaitTime   ' need add BF information
        Else
            AHB_READ_JTAG inAddress, outData, inFieldMask, bDBGlog, dWaitTime
        End If
    Else
        AHB_READDSC inAddress, outData, inFieldMask, bDBGlog, dWaitTime
    End If
    
'      If TheExec.Flow.EnableWord("Open_socket_nWire_Read") = True Then
'        outData = 1
'      End If
      
    Exit Function

ErrHandler:
'   Debug.Print err.Description
'   Stop
'   Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Nwire_Flag_Judgement() As Long
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Nwire_Flag_Judgement"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    g_Nwire_EN = False ''  --> this flag will set to true whec we into function "DSSCtoNwire"
    g_SPMI_EN = False  ''  --> this flag will set to true whec we into function "DSSCtoNwire"
    g_BitField_EN = True '' --> In MP8P always have bitField feature.
    

    'If TheExec.EnableWord("TTR_ALL") = True Then   '' this flag for Module Burst control
    If g_bTTR_ALL = True Then
        KeepAliveStatus = True
        ModBurstStatus = True
    Else
        KeepAliveStatus = False
        ModBurstStatus = False
    End If

    If TheExec.EnableWord("A_Enable_GoNoGo ") = True Then
        GNGStatus = True
    Else
        GNGStatus = False
    End If

    Exit Function
    
ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function SetSpmiEnableFlag(Optional Setflag As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SetSpmiEnableFlag"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    g_SPMI_EN = Setflag
    
    Exit Function
    
ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SetNwireEnableFlag(Optional Setflag As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SetNwireEnableFlag"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    g_Nwire_EN = Setflag
    
    Exit Function
    
ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SetBitFieldEnableFlag(Optional Setflag As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SetBitFieldEnableFlag"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    g_BitField_EN = Setflag
    
    Exit Function
    
ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function GetSpmiEnableFlag() As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetSpmiEnableFlag"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    GetSpmiEnableFlag = g_SPMI_EN
    
    Exit Function
    
ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetNwireEnableFlag() As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetNwireEnableFlag"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    GetNwireEnableFlag = g_Nwire_EN
    
    Exit Function
    
ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetBitFieldEnableFlag() As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetBitFieldEnableFlag"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    GetBitFieldEnableFlag = g_BitField_EN
    
    Exit Function
    
ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function GetModBurstStatus() As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetModBurstStatus"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    GetModBurstStatus = ModBurstStatus
    
    Exit Function
    
ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetKeepAliveStatus() As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetKeepAliveStatus"
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    GetKeepAliveStatus = KeepAliveStatus
    
    Exit Function
    
ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetGNGStatus() As Boolean
    On Error GoTo ErrHandler
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    GetGNGStatus = GNGStatus
    
    Exit Function
    
ErrHandler:
    Debug.Print err.Description
    Stop
    Resume

End Function


' ****************************************** Example : how to use AHB_WRITEDSC *****************************************************
'
' Previous Project (Imola or SStone): AHB_WRITEDSC BUCK0_HP2_CFG_0 , regval
'
' Avus Project : 1. Write Register(Same as previous Project): AHB_WRITEDSC BUCK0_HP2_CFG_0.Addr , regval
'                2. Write Register For Sepcific Field : AHB_WRITEDSC BUCK0_HP2_CFG_0.Addr , regval, BUCK0_HP2_CFG_0.CFG_CC_BLK_SET
'
' For Field Data: Please don't shift the Data!!!!!!!!!!!!!!!!!!!!!!!!
'***********************************************************************************************************************************
'2019/03/27 copy from MP4T/Suzuka, Neo
Public Function AHB_WRITEDSC(Address_In As Variant, ByVal TmpData As SiteLong, Optional Field_Mask_In As Variant = 0) As Long
On Error GoTo ErrHandler
Dim funcName  As String:: funcName = "AHB_WRITEDSC"
Dim dummypat As New PatternSet  'CT 081417
Dim mS_PattArray() As String, mL_PatCount As Long
Dim Data As New SiteLong

Dim Addr As Long
Dim Field_Mask As Long

    '_______________________________________________________________________
    '20200313, for new AHB method
    If TypeName(Address_In) = "String" Then
        Call GetAHB_Add_BF_Value(CStr(Address_In), Addr, Field_Mask)
    Else
        Addr = CLng(Address_In)
        Field_Mask = CLng(Field_Mask_In)
    End If
    '__________________________________________________________________________

Data = TmpData

If TheExec.TesterMode = testModeOffline Then Exit Function '20180910 Need to disable this later

    If TheExec.Sites.Selected.Count = 0 Then
       TheExec.Datalog.WriteComment "*** If no Site alive, bypass the AHB_WRITEDSC **************************************************"
       Exit Function
    End If

    Dim ShiftBits As Long
    Dim num As Integer
    Dim tmpField_Mask As Long
    Dim MaskFiled As Long
    
    If g_sAHB_CHECK = True Then
        'dummypat.Value = g_sAHBBF_WRTIE_TEST_PAT
        dummypat.Value = GetPatListFromPatternSet_OTP(g_sAHBBF_WRTIE_TEST_PAT, mS_PattArray, mL_PatCount)
        TheHdw.Patterns(dummypat).Load
        TheHdw.Patterns(dummypat).Start
        TheHdw.Digital.Patgen.HaltWait
        Exit Function
    End If
    
    'dummypat.Value = g_sAHBBF_WRITE_PAT
    dummypat.Value = GetPatListFromPatternSet_OTP(g_sAHBBF_WRITE_PAT, mS_PattArray, mL_PatCount)
    ShiftBits = 0
    tmpField_Mask = Field_Mask
    For num = 1 To g_iAHB_BW
        If (tmpField_Mask And &H1) = 0 Then
            Exit For
        End If
        tmpField_Mask = Fix(tmpField_Mask / 2)
        ShiftBits = ShiftBits + 1
    Next num
    Data = Data.ShiftLeft(ShiftBits)
    MaskFiled = Field_Mask
    Call Write_32bits(dummypat, g_sTDI, Data, Addr, MaskFiled)

Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'20190416 top
' ****************************************** Example : how to use AHB_WRITEDSC *****************************************************
'
' Previous Project (Imola or SStone): AHB_WRITEDSC BUCK0_HP2_CFG_0 , regval
'
' Avus Project : 1. Write Register(Same as previous Project): AHB_WRITEDSC BUCK0_HP2_CFG_0.Addr , regval
'                2. Write Register For Sepcific Field : AHB_WRITEDSC BUCK0_HP2_CFG_0.Addr , regval, BUCK0_HP2_CFG_0.CFG_CC_BLK_SET
'
' For Field Data: Please don't shift the Data!!!!!!!!!!!!!!!!!!!!!!!!
'***********************************************************************************************************************************
'20180503 evans: modify for field mask method

Public Function AHB_WRITEDSC_NoBitField(Address_In As Variant, ByVal TmpData As SiteLong, Optional Field_Mask_In As Variant = 0, Optional ForceWholeWrite As Boolean = False) As Long
On Error GoTo ErrHandler
Dim funcName  As String:: funcName = "AHB_WRITEDSC_NoBitField"
Dim dummypat As New PatternSet  'CT 081417
Dim mS_PattArray() As String, mL_PatCount As Long
Dim Data As New SiteLong

Dim Addr As Long
Dim Field_Mask As Long

    If TheExec.TesterMode = testModeOffline Then Exit Function '20180910 Need to disable this later

    '_______________________________________________________________________
    '20200313, for new AHB method
    If TypeName(Address_In) = "String" Then
        Call GetAHB_Add_BF_Value(CStr(Address_In), Addr, Field_Mask)
    Else
        Addr = CLng(Address_In)
        Field_Mask = CLng(Field_Mask_In)
    End If
    '_______________________________________________________________________

    Data = TmpData

    If TheExec.Sites.Selected.Count = 0 Then
       TheExec.Datalog.WriteComment "*** If no Site alive, bypass the AHB_WRITEDSC **************************************************"
       Exit Function
    End If

    Dim ShiftBits As Long
    Dim num As Integer
    Dim tmpField_Mask As Long
    Dim MaskFiled As New SiteLong
    
    If g_sAHB_CHECK = True Then
        'dummypat.Value = g_sAHBBF_WRTIE_TEST_PAT
        dummypat.Value = GetPatListFromPatternSet_OTP(g_sAHBBF_WRTIE_TEST_PAT, mS_PattArray, mL_PatCount)
        TheHdw.Patterns(dummypat).Load
        TheHdw.Patterns(dummypat).Start
        TheHdw.Digital.Patgen.HaltWait
        Exit Function
    End If
    
    'dummypat.Value = g_sAHBBF_WRITE_PAT
    
    dummypat.Value = GetPatListFromPatternSet_OTP(g_sAHBBF_WRITE_PAT, mS_PattArray, mL_PatCount)
    
    
    If Field_Mask = 0 Then
        ShiftBits = 0
        Data = Data.ShiftLeft(ShiftBits)
    Else
        Dim outData As New SiteLong
        Dim RealConcern_InData As New SiteLong
        Dim Except_ConcernData_From_ReadData As New SiteLong
    ''''-Read back to combine------------------------------------------------------------------------
        
        If ForceWholeWrite = False Then
            '' Normal setting with Bitfield
            AHB_READDSC Addr, outData, , False
            Except_ConcernData_From_ReadData = outData.BitWiseAnd(Field_Mask)
            ShiftBits = 0
            tmpField_Mask = Field_Mask
            For num = 1 To g_iAHB_BW
                If (tmpField_Mask And &H1) = 0 Then
                    Exit For
                End If
                tmpField_Mask = Fix(tmpField_Mask / 2)
                ShiftBits = ShiftBits + 1
            Next num
            RealConcern_InData = TmpData.ShiftLeft(ShiftBits).BitWiseAnd((2 ^ 8 - 1) - Field_Mask)
            Data = Except_ConcernData_From_ReadData.Add(RealConcern_InData)
        Else
            '' Special setting with Bitfield
            AHB_READDSC Addr, outData, , False
            Except_ConcernData_From_ReadData = outData.BitWiseAnd(Field_Mask)
            ShiftBits = 0
            RealConcern_InData = TmpData.ShiftLeft(ShiftBits).BitWiseAnd((2 ^ 8 - 1) - Field_Mask)
            Data = Except_ConcernData_From_ReadData.Add(RealConcern_InData)
        End If
        
    End If
    
    Call Write_24bits(dummypat, g_sTDI, Data, Addr)

Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'20180503 evans: create for field mask method
Public Function Write_32bits(TrimPattern As PatternSet, JTAG_TDI As String, Data As SiteLong, Addr As Long, Field_Mask As Long) As Long
On Error GoTo ErrHandler
Dim funcName  As String:: funcName = "Write_32bits"
    Dim addrSerial As New DSPWave
    Dim dataSerial As New DSPWave
    Dim bfSerial As New DSPWave
    Dim addressplusdataWave As New DSPWave
    Dim SignalName As String
    Dim WaveDef As String
    Dim addrwidth As Long
    Dim DataWidth As Long
    Dim BFWidth As Long
    Dim i As Long
    Dim dataout As New SiteLong
    Dim Address As New SiteLong
        
    addrwidth = g_iAHB_ADDR_BW
    DataWidth = g_iAHB_BW
    BFWidth = g_iAHB_DATA_BW

    For Each Site In TheExec.Sites
        addrSerial.CreateConstant Addr, 1, DspLong
        dataSerial.CreateConstant Data, 1, DspLong
        bfSerial.CreateConstant Field_Mask, 1, DspLong
    Next Site
 
    For Each Site In TheExec.Sites
        bfSerial = bfSerial.ConvertStreamTo(tldspSerial, DataWidth, 0, Bit0IsMsb)
        dataSerial = dataSerial.ConvertStreamTo(tldspSerial, DataWidth, 0, Bit0IsMsb)
        addrSerial = addrSerial.ConvertStreamTo(tldspSerial, addrwidth, 0, Bit0IsMsb)
        addressplusdataWave = bfSerial.Copy
        addressplusdataWave = addressplusdataWave.Concatenate(dataSerial).repeat(1)
        addressplusdataWave = addressplusdataWave.Concatenate(addrSerial).repeat(1)
    Next Site
    
    WaveDef = "WaveDef"
    SignalName = "Addressplusdata"
    
    TheHdw.Patterns(TrimPattern).Load
    For Each Site In TheExec.Sites
        TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & Site, addressplusdataWave, True
        TheHdw.DSSC.Pins(JTAG_TDI).Pattern(TrimPattern).Source.Signals.Add SignalName
        With TheHdw.DSSC.Pins(JTAG_TDI).Pattern(TrimPattern).Source.Signals(SignalName)
            .WaveDefinitionName = WaveDef & Site
            .SampleSize = (addrwidth + DataWidth + BFWidth)
            .Amplitude = 1
            '.LoadSamples
            .LoadSettings
        End With
    Next Site
    
    TheHdw.DSSC.Pins(JTAG_TDI).Pattern(TrimPattern).Source.Signals.DefaultSignal = SignalName
    
    TheHdw.Patterns(TrimPattern).Start ("")
    
    ' Bypass DSP computing, use HOST computer
    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
    ' Halt on opcode to make sure all samples are capture.
    TheHdw.Digital.Patgen.HaltWait
    
    Exit Function
   
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Write_24bits(TrimPattern As PatternSet, JTAG_TDI As String, Data As SiteLong, Addr As Long) As Long
On Error GoTo ErrHandler
Dim funcName  As String:: funcName = "Write_24bits"
    Dim addrSerial As New DSPWave
    Dim dataSerial As New DSPWave
    Dim addressplusdataWave As New DSPWave
    Dim SignalName As String
    Dim WaveDef As String
    Dim addrwidth As Long
    Dim DataWidth As Long
    Dim i As Long
    Dim dataout As New SiteLong
    Dim Address As New SiteLong
    
'    addrwidth = 16
'    DataWidth = 8
    addrwidth = g_iAHB_ADDR_BW
    DataWidth = g_iAHB_BW

    For Each Site In TheExec.Sites
    addrSerial.CreateConstant Addr, 1, DspLong
    dataSerial.CreateConstant Data, 1, DspLong
    Next Site
 
     For Each Site In TheExec.Sites
        dataSerial = dataSerial.ConvertStreamTo(tldspSerial, DataWidth, 0, Bit0IsMsb)
        addrSerial = addrSerial.ConvertStreamTo(tldspSerial, addrwidth, 0, Bit0IsMsb)
        addressplusdataWave = dataSerial.Copy
        addressplusdataWave = addressplusdataWave.Concatenate(addrSerial).repeat(1)
     Next Site

    
    WaveDef = "WaveDef"
    SignalName = "Addressplusdata"
    
    TheHdw.Patterns(TrimPattern).Load
    For Each Site In TheExec.Sites
        TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & Site, addressplusdataWave, True
        TheHdw.DSSC.Pins(JTAG_TDI).Pattern(TrimPattern).Source.Signals.Add SignalName
        With TheHdw.DSSC.Pins(JTAG_TDI).Pattern(TrimPattern).Source.Signals(SignalName)
            .WaveDefinitionName = WaveDef & Site
            .SampleSize = (addrwidth + DataWidth)
            .Amplitude = 1
            '.LoadSamples
            .LoadSettings
        End With
    Next Site
    
    TheHdw.DSSC.Pins(JTAG_TDI).Pattern(TrimPattern).Source.Signals.DefaultSignal = SignalName
    
    
    TheHdw.Patterns(TrimPattern).Start ("")
    
    
    ' Bypass DSP computing, use HOST computer
    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
    ' Halt on opcode to make sure all samples are capture.
    TheHdw.Digital.Patgen.HaltWait
    
   Exit Function
   
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


' ****************************************** Example : how to use AHB_READDSC ************************************************
'
' Previous Project (Imola or SStone): AHB_READDSC BUCK0_HP2_CFG_0 , regval
'
' Avus Project : 1. Read Register(Same as previous Project): AHB_READDSC BUCK0_HP2_CFG_0.Addr , regval
'                2. Read Register By Field : AHB_READDSC BUCK0_HP2_CFG_0.Addr , regval, BUCK0_HP2_CFG_0.CFG_CC_BLK_SET
'
'*****************************************************************************************************************************
'2019/03/27 copy from MP4T/Suzuka, Neo
Public Function AHB_READDSC(Address_In As Variant, Data As SiteLong, Optional Field_Mask_In As Variant = 0, Optional bDBGlog As Boolean = True, Optional dWaitTime As Double = 10 * ms) As SiteLong
On Error GoTo ErrHandler
Dim funcName  As String:: funcName = "AHB_READDSC"
Dim dummypat        As New PatternSet       'CT 081417
Dim mS_PattArray()  As String, mL_PatCount      As Long
Dim BitAND          As Long, Offset             As Long
Dim BitANDStr       As String, CalcData         As New SiteLong
Dim Site As Variant
Dim Address As Long
Dim Field_Mask As Long

    If TheExec.TesterMode = testModeOffline Then Exit Function '20180910 Need to disable later

    If TheExec.Sites.Selected.Count = 0 Then
       TheExec.Datalog.WriteComment "*** If no Site alive, bypass the AHB_READDSC **************************************************"
       Exit Function
    End If
    
    '_______________________________________________________________________
    '20200313, for new AHB method
    If TypeName(Address_In) = "String" Then
        Call GetAHB_Add_BF_Value(CStr(Address_In), Address, Field_Mask)
    Else
        Address = CLng(Address_In)
        Field_Mask = CLng(Field_Mask_In)
    End If
    '_______________________________________________________________________
    
    Address = Address And &HFFFF&
    dummypat.Value = GetPatListFromPatternSet_OTP(g_sAHB_READ_PAT, mS_PattArray, mL_PatCount)
    
    
    Data_read dummypat, g_sTDI, g_sTDO, CLng(Address), Data, bDBGlog, dWaitTime

    If Field_Mask > 0 Then
        BitAND = (&HFF) Xor Field_Mask
        BitANDStr = ConvertFormat_Dec2Bin_Complement(BitAND, 8) 'was FormatConv_Dec2Bin_Complement
        Offset = InStr(StrReverse(BitANDStr), "1") - 1
        CalcData = Data.BitWiseAnd(BitAND).ShiftRight(Offset)
        Data = CalcData
        For Each Site In TheExec.Sites
                        'Buck6 needs data print out in the datalog file.  Please use Top level "False" if debug print is not needed.
            If bDBGlog = True Then TheExec.Datalog.WriteComment "Adress-h'" & Hex(Address) & "(d'" & (Address And &HFF) & ")/" & "Data-" & Hex(Data(Site))
            If bDBGlog = True Then Debug.Print "Adress-h'" & Hex(Address) & "(d'" & (Address And &HFF) & ")/" & "Data-" & Hex(Data(Site))
        Next Site
    End If

Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___20200313 for AHB New Method by OTPData Structure
Public Function GetAHB_Add_BF_Value(Address As String, Address_out As Long, FieldMask_out As Long)
On Error GoTo ErrHandler
Dim funcName  As String:: funcName = "GetAHB_Add_BF_Value"
Dim AHB_Idx As String

    AHB_Idx = g_dictAHBEnumIdx.Item(UCase(Address))
    Address_out = CLng(g_OTPData.Category(AHB_Idx).sAhbAddress)
    FieldMask_out = CLng(g_OTPData.Category(AHB_Idx).lAhbMaskVal)
    If InStr(Address, ".") = False Then FieldMask_out = 0
    
Exit Function

ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___20200313, New from MP7P
Public Function Get_AHBBF2Long(BF_Name As String) As Long
On Error GoTo ErrHandler
Dim funcName  As String:: funcName = "Get_AHBBF2Long"
Dim AHB_Idx As String

    AHB_Idx = g_dictAHBEnumIdx.Item(UCase(BF_Name))
    Get_AHBBF2Long = CLng(g_OTPData.Category(AHB_Idx).lAhbMaskVal) 'was .AHB_Mask_Value

Exit Function

ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___20200313, New from MP7P
Public Function Get_AHBAddr2Long(Address As String) As Long
On Error GoTo ErrHandler
Dim funcName  As String:: funcName = "Get_AHBAddr2Long"
Dim AHB_Idx As String

    AHB_Idx = g_dictAHBEnumIdx.Item(UCase(Address))
    Get_AHBAddr2Long = CLng(g_OTPData.Category(AHB_Idx).sAhbAddress) 'was .AHB_Addr

Exit Function

ErrHandler:
'    Debug.Print err.Description
'    Stop
'    Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___20200313, Move from module VBT_LIB_Common
Public Function Data_read(PatName As PatternSet, JTAG_TDI As String, JTAG_TDO As String, Addr As Long, ByRef Data As SiteLong, Optional bDBGlog As Boolean = True, Optional dWaitTime As Double = 10 * ms)
    'here we are using digsource and digcap at the same time so we write to the adress and read from the adress at the same time
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Data_read"

    Dim addrwidth As Long
    Dim DataWidth As Long

    Dim i         As Long
    Dim dataout   As New SiteLong
    Dim Address   As New SiteLong
    Dim Site      As Variant

    If TheExec.Sites.Active.Count = 0 Then Exit Function

    Address = ToSiteLong(Addr)

    addrwidth = 16
    DataWidth = 16

    Dim addressSerial As New DSPWave, AdressWave As New DSPWave

    For Each Site In TheExec.Sites
        addressSerial.CreateConstant Address, 1, DspLong
        addressSerial = addressSerial.ConvertStreamTo(tldspSerial, addrwidth, 0, Bit0IsMsb)
        AdressWave = addressSerial.Copy.repeat(2)
        'AdressWave = addressSerial.Copy.repeat(1) '20171118
    Next Site

    Dim SignalName As String
    Dim WaveDef   As String
    WaveDef = "Wavedef"
    SignalName = "OnlyAdress"

    addrwidth = 32

    TheHdw.Patterns(PatName).Load

    For Each Site In TheExec.Sites.Selected
        TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & Site, AdressWave, True

        TheHdw.DSSC.Pins(JTAG_TDI).Pattern(PatName).Source.Signals.Add SignalName
        With TheHdw.DSSC(JTAG_TDI).Pattern(PatName).Source.Signals(SignalName)
            .Reinitialize
            .WaveDefinitionName = WaveDef & Site
            .Amplitude = 1
            .SampleSize = addrwidth    'dressWave.SampleSize 'addrwidth '20171118
            '.LoadSamples
            .LoadSettings

        End With
    Next Site

    TheHdw.DSSC(JTAG_TDI).Pattern(PatName).Source.Signals.DefaultSignal = SignalName

    'setup capture
    Dim DataArray As New DSPWave
    DataWidth = 8
    DataArray.CreateConstant 0, 8
    Call DSSC_Capture_Setup(PatName, JTAG_TDO, "dataSigAHB", DataWidth, DataArray, dWaitTime)

    ' Bypass DSP computing, use HOST computer
    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
    ' Halt on opcode to make sure all samples are capture.
    TheHdw.Digital.Patgen.HaltWait

    For Each Site In TheExec.Sites
        If DataArray.SampleSize <> 8 Then
            TheExec.Datalog.WriteComment "DataArray.SampleSize=" & CStr(DataArray.SampleSize)    '20171118
            'Stop
        End If
        Data(Site) = (DataArray.Element(7) * 2 ^ 7 + DataArray.Element(6) * 2 ^ 6 + DataArray.Element(5) * 2 ^ 5 + DataArray.Element(4) * 2 ^ 4 + DataArray.Element(3) * 2 ^ 3 + DataArray.Element(2) * 2 ^ 2 + DataArray.Element(1) * 2 ^ 1 + DataArray.Element(0) * 2 ^ 0)
    Next Site

    For Each Site In TheExec.Sites
        'Buck6 needs data print out in the datalog file.  Please use Top level "False" if debug print is not needed.
        If bDBGlog = True Then TheExec.Datalog.WriteComment "Adress-h'" & Hex(Addr) & "(d'" & (Addr And &HFF) & ")/" & "Data-" & Hex(Data(Site))
        If bDBGlog = True Then Debug.Print "Adress-h'" & Hex(Addr) & "(d'" & (Addr And &HFF) & ")/" & "Data-" & Hex(Data(Site))
    Next Site

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'___20200313, Move from module VBT_LIB_Common
'******************************************************************************
'' Digital Signal Capture utilities
''******************************************************************************
'2019/03/27 copy from MP4T/Suzuka, Neo
Public Function DSSC_Capture_Setup(PatName As PatternSet, DigCapPin As String, _
                SignalName As String, SampleSize As Long, CapWave As DSPWave, Optional dWaitTime As Double = 10 * ms)
    
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "DSSC_Capture_Setup"

    'TheHdw.Patterns(PatName).Load
    With TheHdw.DSSC.Pins(DigCapPin).Pattern(PatName).Capture.Signals
        .Reinitialize
        .Add (SignalName)
        With .Item(SignalName)
            .Reinitialize
            .SampleSize = SampleSize
            .LoadSettings
        End With
    End With
    
    ' Bind capture results to DSPWave object
    'CapWave = TheHdw.DSSC.Pins(DigCapPin).Pattern(PatName).Capture.Signals(SignalName).DSPWave   'WAS.  20171118 REMOVE
    
    'Bypass DSP computing, use HOST computer '20171118
    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
    'halt on opcode to make sure all samples are capture.
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    ' Bind capture results to DSPWave object
  'CapWave = TheHdw.DSSC.Pins(DigCapPin).Pattern(PatName).Capture.Signals(SignalName).DSPWave
    TheHdw.Wait dWaitTime ' was fixed 10 * ms 'add wait time for acore RPoly open kelvin alarm
    
    ''TheHdw.Patterns(PatName).start ("")
    Call TheHdw.Patterns(PatName).test(pfNever, 0)
    ' TheHdw.Wait 2

    TheHdw.Digital.Patgen.HaltWait

    ' Bind capture results to DSPWave object
    CapWave = TheHdw.DSSC.Pins(DigCapPin).Pattern(PatName).Capture.Signals(SignalName).DSPWave   '20171118

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190905 TH
Public Function nWire_Setup()

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "nWire_Setup"
    
    If TheExec.Sites.Selected.Count = 0 Then Exit Function

    'init to set g_BitField_EN as true (depends on project design)

    g_BitField_EN = True

    If UCase(TheExec.DataManager.InstanceName) Like "*2NWIRESPMI*" Then
        '' need to be careful to control SPMI enable, below setting dependent on different project setting naming
        ''g_RegVal = &H10: AHB_WRITEDSC HOST_INTERFACE_SPMI_REGISTERS_SPMI_CTRL.Addr, g_RegVal
        
        '___20200313, for AHB New Method
        g_RegVal = &H10: AHB_WRITEDSC "HOST_INTERFACE_SPMI_REGISTERS_SPMI_CTRL", g_RegVal

        TheHdw.Protocol.ports("NWIRE_SPMI").Enabled = True
        TheHdw.Protocol.ports("nWire_SPMI").NWire.HRAM.Setup.TriggerType = tlNWireHRAMTriggerType_Never
        TheHdw.Protocol.ports("nWire_SPMI").NWire.HRAM.Setup.WaitForEvent = False
        TheHdw.Protocol.ModuleRecordingEnabled = True    '?
        'Call ENABLE_SPMI_PA
        g_Nwire_EN = True
        g_SPMI_EN = True

        '___20200313, for AHB New Method
'        Call AHB_WRITE(FABRIC_AHB_FABRIC_RMWU_MODE.Addr, ToSiteLong(&H2))
        Call AHB_WRITE("FABRIC_AHB_FABRIC_RMWU_MODE", ToSiteLong(&H2))
'        Call AHB_WRITE(FABRIC_AHB_FABRIC_RMWU_MAIN_EN.Addr, ToSiteLong(&H1))
        Call AHB_WRITE("FABRIC_AHB_FABRIC_RMWU_MAIN_EN", ToSiteLong(&H1))
        
        TheExec.Datalog.WriteComment "SWITCH to SPMI PA!!!"

    ElseIf UCase(TheExec.DataManager.InstanceName) Like "*2NWIREJTAG*" Then
        TheHdw.Protocol.ports("NWIRE_JTAG").Enabled = True
        TheHdw.Protocol.ports("NWIRE_JTAG").NWire.HRAM.Setup.TriggerType = tlNWireHRAMTriggerType_Never
        TheHdw.Protocol.ports("NWIRE_JTAG").NWire.HRAM.Setup.WaitForEvent = False
        g_Nwire_EN = True
        g_SPMI_EN = False

        TheExec.Datalog.WriteComment "SWITCH to JTAG PA!!!"

    ElseIf UCase(TheExec.DataManager.InstanceName) Like "*NWIRESPMI2*" Then
        TheHdw.Protocol.ports("NWIRE_SPMI").Halt
        TheHdw.Protocol.ports("NWIRE_SPMI").Enabled = False
        g_Nwire_EN = False
        g_SPMI_EN = False

        TheExec.Datalog.WriteComment "SWITCH from SPMI_PA to DSSC!!!"

    ElseIf UCase(TheExec.DataManager.InstanceName) Like "*NWIREJTAG2*" Then
        TheHdw.Protocol.ports("NWIRE_JTAG").Halt
        TheHdw.Protocol.ports("NWIRE_JTAG").Enabled = False
        g_Nwire_EN = False
        g_SPMI_EN = False

        TheExec.Datalog.WriteComment "SWITCH from JTAG_PA to DSSC!!!"
    End If

    Exit Function
ErrHandler:
    'Debug.Print err.Description
    '    Stop   '//2019_1213
    'Resume
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

'___________________________________________________________________________________________________
'___20200313 Unused Codes
'___________________________________________________________________________________________________
''
'''20171120 evans add mask for trim dsc
''Public Function AHB_WRITENWIRE_TrimWithMask(Address As Long, Data As SiteLong, mask As SiteLong, Optional Offset As Long = 0, Optional maskoffset As Long = 0) As Long
''On Error GoTo ErrHandler
''
''    Dim PortName As String
''    PortName = "NWIRE_JTAG"
''
''    Dim cap_SDOVAL As SiteLong
''
''    Dim myPLD As New PinListData
''    Dim DatafromPA As New SiteLong
''    Dim PariyfromPA As New SiteLong
''    Dim CMEMtransindexes As INWireCMEMTransactionIndexes
''    Dim TransIndex As INWireCMEMTransactionIndex
''
''    Dim nsite As Variant
''    Dim TestName As String
''    Dim Address_Data As New SiteLong
''
'''    For Each g_Site In TheExec.Sites.Selected
'''        Address_Data(g_Site) = AddressVal * 2 ^ 8 + DataVal(g_Site)
'''    Next g_Site
''
''
''''    '//If All sites valure are same then execute WriteALL = 1 with Recording
''''    If WriteAll = 1 Then
''''
''''        TestName = TheExec.DataManager.InstanceName
''''        For Each g_Site In TheExec.Sites.Selected
''''            TestName = TestName & CStr(Hex(Address(g_Site)))
''''        Next g_Site
''''
''''        If thehdw.Protocol.Ports(PortName).Modules.IsRecorded(TestName) = False Then
''''
''''            With thehdw.Protocol.Ports(PortName).NWire.Frames("WriteIR2")
''''                .Fields("Mask").Value = 0 ' &H4E1302
''''                .Fields("Addr").Value = Address ' &H4E1302
''''                .Fields("Data").Value = Data ' &H4E1302
''''                .Execute tlNWireExecutionType_Default
''''            End With
''''
''''            thehdw.Protocol.Ports(PortName).IdleWait
''''            thehdw.Protocol.Ports(PortName).Modules.StopRecording
''''        End If
''''    '//If All sites valure are different then execute WriteALL = 2 without Recording
''''    ElseIf WriteAll = 2 Then
''            With TheHdw.Protocol.ports(PortName).NWire.Frames("Write")
''                .Fields("Mask").Value = 0 ' &H4E1302
''                .Fields("Addr").Value = Address ' &H4E1302
''                .Fields("Data").Value = Data ' &H4E1302
''                .Execute tlNWireExecutionType_Default
''            End With
''            TheHdw.Protocol.ports(PortName).IdleWait
''''    End If
''
''Exit Function
''
''ErrHandler:
''   Debug.Print err.Description
''   Stop
''   Resume
''
''
''End Function

'''---------------------------------------------------------------------------------------------------------------------------------
'''---------------------------------------------------------------------------------------------------------------------------------
'''---------------------------------------------------------------------------------------------------------------------------------
''
''
''Public Function Write_24bits_NWIRE(TrimPattern As PatternSet, JTAG_TDI As String, Data As SiteLong, Addr As SiteLong) As Long
''
''    Dim addrSerial As New DSPWave
''    Dim dataSerial As New DSPWave
''    Dim addressplusdataWave As New DSPWave
''    Dim SignalName As String
''    Dim WaveDef As String
''    Dim addrwidth As Long
''    Dim DataWidth As Long
''    Dim i As Long
''    Dim dataout As New SiteLong
''    Dim Address As New SiteLong
''
''    addrwidth = 16
''    DataWidth = 8
''
''
''    On Error GoTo ErrHandler
''
''
''    For Each Site In TheExec.Sites
''    addrSerial.CreateConstant Addr, 1, DspLong
''    dataSerial.CreateConstant Data, 1, DspLong
''    Next Site
''
''     For Each Site In TheExec.Sites
''        dataSerial = dataSerial.ConvertStreamTo(tldspSerial, DataWidth, 0, Bit0IsMsb)
''        addrSerial = addrSerial.ConvertStreamTo(tldspSerial, addrwidth, 0, Bit0IsMsb)
''        addressplusdataWave = dataSerial.Copy
''        addressplusdataWave = addressplusdataWave.Concatenate(addrSerial).repeat(1)
''     Next Site
''
''
''    WaveDef = "WaveDef"
''    SignalName = "Addressplusdata"
''
''    TheHdw.Patterns(TrimPattern).Load
''    For Each Site In TheExec.Sites
''        TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & Site, addressplusdataWave, True
''        TheHdw.DSSC.Pins(JTAG_TDI).Pattern(TrimPattern).Source.Signals.Add SignalName
''        With TheHdw.DSSC.Pins(JTAG_TDI).Pattern(TrimPattern).Source.Signals(SignalName)
''            .WaveDefinitionName = WaveDef & Site
''            .SampleSize = (addrwidth + DataWidth)
''            .Amplitude = 1
''            .LoadSamples
''            .LoadSettings
''        End With
''    Next Site
''
''    TheHdw.DSSC.Pins(JTAG_TDI).Pattern(TrimPattern).Source.Signals.DefaultSignal = SignalName
''
''
''    TheHdw.Patterns(TrimPattern).Start ("")
''
''
''    ' Bypass DSP computing, use HOST computer
''    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
''    ' Halt on opcode to make sure all samples are capture.
''    TheHdw.Digital.Patgen.HaltWait
''
''
''
''
''
''   Exit Function
''
''ErrHandler:
''    If AbortTest Then Exit Function Else Resume Next
''
''End Function

'''---------------------------------------------------------------------------------------------------------------------------------
'''---------------------------------------------------------------------------------------------------------------------------------
'''---------------------------------------------------------------------------------------------------------------------------------


'''
'''' ****************************************** Example : how to use AHB_READDSC ************************************************
''''
'''' Previous Project (Imola or SStone): AHB_READDSC BUCK0_HP2_CFG_0 , regval
''''
'''' Avus Project : 1. Read Register(Same as previous Project): AHB_READDSC BUCK0_HP2_CFG_0.Addr , regval
''''                2. Read Register By Field : AHB_READDSC BUCK0_HP2_CFG_0.Addr , regval, BUCK0_HP2_CFG_0.CFG_CC_BLK_SET
''''
''''*****************************************************************************************************************************
''''2018/05/26
'''Public Function AHB_READNWIRE_ISENSE(inAddress As Long, outData As SiteLong, Optional Field_Mask As Long = 0, Optional bDBGlog As Boolean = True, Optional dWaitTime As Double = 10 * ms) As SiteLong
'''    On Error GoTo ErrHandler
'''
'''
'''    Dim PortName As String
'''    PortName = "NWIRE_JTAG"
'''
'''    Dim myPLD_nWire As New PinListData
''''    Dim indexes(0 To 1) As INWireHRAMTransactionIndexes
'''    Dim indexes_CMEM_0 As INWireCMEMTransactionIndexes
'''    Dim indexes_CMEM_1 As INWireCMEMTransactionIndexes
'''
'''    Dim TempData As New SiteDouble
'''
'''    If TheExec.TesterMode = testModeOffline Then Exit Function '20181017
'''
'''STT
'''    With TheHdw.Protocol.ports(PortName).NWire.Frames("Read")
'''        .Fields("Addr").Value = 32259 'GPADC_ADC_CFG_MANUAL_1_CONV_RES_MSB
'''        .Execute tlNWireExecutionType_CaptureInCMEM
'''        TheHdw.Protocol.ports(PortName).IdleWait
'''
'''        For Each g_Site In TheExec.Sites
'''            Set indexes_CMEM_0 = TheHdw.Protocol.ports(PortName).NWire.CMEM.Transactions.Read(0, 0).Pins(PortName).Value
'''        Next g_Site
'''
'''        .Fields("Addr").Value = 32258 'GPADC_ADC_CFG_MANUAL_1_CONV_RES_LSB
'''        .Execute tlNWireExecutionType_CaptureInCMEM
'''        TheHdw.Protocol.ports(PortName).IdleWait
'''
'''        For Each g_Site In TheExec.Sites
'''            Set indexes_CMEM_1 = TheHdw.Protocol.ports(PortName).NWire.CMEM.Transactions.Read(0, 1).Pins(PortName).Value
'''        Next g_Site
'''
'''        outData = indexes_CMEM_0(0).Fields("Data").Value
'''        outData = indexes_CMEM_1(0).Fields("Data").Value
'''
'''
'''    End With
'''
'''
'''SPT
'''
'''
'''
'''    TheHdw.Protocol.ports(PortName).IdleWait
'''
'''    TheHdw.Wait 0
'''
'''    '******************************************************************************************
'''    Dim BitAND          As Long, Offset             As Long
'''    Dim BitANDStr       As String, CalcData         As New SiteLong
'''
'''    If Field_Mask > 0 Then
'''        BitAND = (&HFF) Xor Field_Mask
'''        BitANDStr = auto_Dec2Bin_OTP(BitAND, 8)
'''        Offset = InStr(StrReverse(BitANDStr), "1") - 1
'''        CalcData = outData.BitwiseAnd(BitAND).ShiftRight(Offset)
'''        outData = CalcData
'''    End If
'''
'''
'''    For Each Site In TheExec.Sites
'''        'If bDBGlog = True Then TheExec.Datalog.WriteComment "Address-h'" & Hex(inAddress) & "(d'" & (inAddress And &HFF) & ")/" & "Data-" & Hex(outData(Site))
''''        If bDBGlog = True Then Debug.Print "Address-h'" & Hex(inAddress) & "(d'" & (inAddress And &HFF) & ")/" & "Data-" & Hex(outData(Site))
'''    Next Site
'''
'''''    For Each g_Site In TheExec.Sites.Selected
'''''        Debug.Print "g_Site:" & g_Site & " Address-h'" & Hex(inAddress) & "/" & "Data-h'" & Hex(outData(g_Site))
'''''    Next g_Site
'''
'''Exit Function
'''
'''ErrHandler:
'''        Debug.Print err.Description
'''        Stop
'''    Resume
'''End Function


