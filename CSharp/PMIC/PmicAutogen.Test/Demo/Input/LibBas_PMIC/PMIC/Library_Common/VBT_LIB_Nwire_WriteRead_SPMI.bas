Attribute VB_Name = "VBT_LIB_Nwire_WriteRead_SPMI"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\simetra\Simetra_A0_TestPlan_190731_Jeremy.xlsx, MD5=f21f102c34ba8020743b845115cd9abb
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

Private bSPMI_PA  As Boolean
Global Const pcPAMask = True
Private bSPMIEnabled As Boolean

Private Const cParity = 1    'not 1: 0dd, not 0: even
Private Const BIT0 = 1

Public Enum eSPMIStatus
    eEnable = True
    eDisable = False
End Enum

'Public Function SPMI_CHECK()
'    ENABLE_SPMI_PA
'Exit Function
'''''
'''''
'''''    RegVal = &H1
'''''    Call AHB_WRITEDSC(BUCK4_HP1_CFG_2.Addr, RegVal, BUCK4_HP1_CFG_2.BUCK4_HP1_CFG_2_CFG_CS_GAIN_TRIM)
'''''    g_RegVal = &H3C: AHB_WRITEDSC GPADC_ADC_CFG_MUXCTRL_OVERRIDE_VAL0.Addr, g_RegVal
'''''    RegVal = &H0
'''''    Call AHB_READDSC(BUCK4_HP1_CFG_2.Addr, RegVal)
''''''''    Call TheHdw.Digital.Pins("SPMI_SCLK,SPMI_SDATA").Disconnect
''''''''    Call TheHdw.Digital.Pins("SPMI_SCLK_PA,SPMI_SDATA_PA").Connect
''''''''    TheHdw.StartStopwatch
''''''''    RegVal = &HA
''''''''    Call SPMI_PA_WRITE(GPADC_ADC_CFG_ATB_OVERRIDE.Addr, RegVal)
''''''''    TheExec.Datalog.WriteComment "SPMI_PA_WRITE:" & TheHdw.ReadStopwatch
''''''''
''''''''    TheHdw.StartStopwatch
''''''''    RegVal = &H0
''''''''    Call SPMI_PA_READ(GPADC_ADC_CFG_ATB_OVERRIDE.Addr, RegVal)
''''''''    TheExec.Datalog.WriteComment "SPMI_PA_READ:" & TheHdw.ReadStopwatch
'''''Stop
'''''
'''''    Call TheHdw.Digital.Pins("SPMI_SCLK,SPMI_SDATA").Connect
'''''    Call TheHdw.Digital.Pins("SPMI_SCLK_PA,SPMI_SDATA_PA").Disconnect
'''''    RegVal = &H0
'''''    Call SPMI_READ(GPADC_ADC_CFG_ATB_OVERRIDE.Addr, RegVal)
'''''''''''    TheExec.Datalog.WriteComment "SPMI_READ:" & TheHdw.ReadStopwatch
''''''''Stop
'''''''''''    TheHdw.StartStopwatch
'''''    RegVal = &H9
'''''    Call SPMI_WRITE(GPADC_ADC_CFG_ATB_OVERRIDE.Addr, RegVal)
''''''''    TheExec.Datalog.WriteComment "SPMI_WRITE:" & TheHdw.ReadStopwatch
''''''Stop
''''''''    TheHdw.StartStopwatch
'''''    RegVal = &H0
'''''    Call AHB_READDSC(GPADC_ADC_CFG_ATB_OVERRIDE.Addr, RegVal)
'''''Stop
''''''''    TheExec.Datalog.WriteComment "AHB_READDSC:" & TheHdw.ReadStopwatch
'''''
''''''''    TheHdw.StartStopwatch
'''''    RegVal = &H3
'''''    Call AHB_WRITEDSC(GPADC_ADC_CFG_ATB_OVERRIDE.Addr, RegVal)
''''''''    TheExec.Datalog.WriteComment "AHB_WRITEDSC:" & TheHdw.ReadStopwatch
'''''    Stop


'Public Sub SPMI_PA_Initialize()
'
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "SPMI_PA_Initialize"
'
'    bSPMI_PA = False
'    bSPMIEnabled = False
'
'    Exit Sub
'
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Sub Else Resume Next
'End Sub

'Public Function Get_SPMI_STATUS() As eSPMIStatus
'
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "Get_SPMI_STATUS"
'
'    If bSPMIEnabled = False Then
'        Get_SPMI_STATUS = eDisable
'    ElseIf bSPMIEnabled = True Then
'        If bSPMI_PA = True Then
'            Get_SPMI_STATUS = eEnable
'        ElseIf bSPMI_PA = False Then
'            Get_SPMI_STATUS = eDisable
'        End If
'    End If
'
'    Exit Function
'
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'Public Function ENABLE_SPMI_PA()
'
''' -------------------------------------------------------------------------------------------
''' ----Need to check register setting for SPMI , different device have different setting------
''' -------------------------------------------------------------------------------------------
'    Dim PortName As String
'    PortName = "NWIRE_SPMI"
'
'
'    #If 0 Then
'        RegVal = &H10: AHB_WRITEDSC HOST_INTERFACE_SPMI_REGISTERS_SPMI_CTRL.Addr, RegVal
'        Call AHB_WRITEDSC(FABRIC_AHB_FABRIC_RMWU_MODE.Addr, ToSiteLong(&H2))
'        Call AHB_WRITEDSC(FABRIC_AHB_FABRIC_RMWU_MAIN_EN.Addr, ToSiteLong(&H1))
'
'    '    Call TheHdw.Digital.Pins("SPMI_SCLK,SPMI_SDATA").Disconnect
'        Call TheHdw.Digital.Pins("SPMI_SCLK,SPMI_SDATA").Connect
'
'        TheHdw.Protocol.ports(PortName).Enabled = True
'        TheHdw.Protocol.ports(PortName).NWire.HRAM.Setup.TriggerType = tlNWireHRAMTriggerType_Never
'        TheHdw.Protocol.ports(PortName).NWire.HRAM.Setup.WaitForEvent = False
'        TheHdw.Protocol.ModuleRecordingEnabled = True
'
'        bSPMI_PA = True
'        bSPMIEnabled = True
'    '    Call AHB_WRITEDSC(FABRIC_AHB_FABRIC_RMWU_MODE.Addr, ToSiteLong(&H2))
'
'        TheExec.Datalog.WriteComment "ENABLE and SWITCH to use SPMI PA!!!"
'    #End If
'
'End Function

'Public Function DSSC_2_SPMIPA_Switch()
'
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "DSSC_2_SPMIPA_Switch"
'
'    If bSPMIEnabled = True Then
'        bSPMI_PA = True
'        TheExec.Datalog.WriteComment "SWITCH to use SPMI PA!!!"
'        'g_RegVal = &H2: AHB_WRITEDSC FABRIC_AHB_FABRIC_RMWU_MODE.Addr, g_RegVal
'        '___20200313, New AHB Method
'        g_RegVal = &H2: AHB_WRITEDSC "FABRIC_AHB_FABRIC_RMWU_MODE", g_RegVal
'    Else
'        TheExec.Datalog.WriteComment "Can't be switch because SPMI PA is not enabled!!!"
'    End If
'
'    Exit Function
'
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function
'
'Public Function SPMIPA_2_DSSC_Switch()
'
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "SPMIPA_2_DSSC_Switch"
'
'    If bSPMIEnabled = True Then
'        bSPMI_PA = False
'        TheExec.Datalog.WriteComment "SWITCH to use AHB DSSC!!!"
'    Else
'        TheExec.Datalog.WriteComment "Can't be switch because SPMI PA is not enabled!!!"
'    End If
'
'    Exit Function
'
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function

'Public Sub SPMIBF_PA_WRITE(Addr As Long, TmpData As SiteLong, Optional FieldMsk As Long = 0)
'___20200313, New AHB Method
Public Sub SPMIBF_PA_WRITE(Address_In As Variant, TmpData As SiteLong, Optional Field_Mask_In As Variant = 0)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SPMIBF_PA_WRITE"

    Dim PortName  As String
    PortName = "NWIRE_SPMI"

    Dim Address   As Long
    Dim Data      As New SiteLong
    Dim BitField  As Long

    Dim ShiftBits As Long
    Dim tmpField_Mask As Long
    Dim num As Integer

    '_______________________________________________________________________
    '20200313, for new AHB method
    Dim Addr As Long
    Dim FieldMsk As Long
    If TypeName(Address_In) = "String" Then
        Call GetAHB_Add_BF_Value(CStr(Address_In), Addr, FieldMsk)
    Else
        Addr = CLng(Address_In)
        FieldMsk = CLng(Field_Mask_In)
    End If
    '_______________________________________________________________________
    
    '___bit shift according to BF
    ShiftBits = 0
    tmpField_Mask = FieldMsk


    For num = 1 To g_iAHB_BW
        If (tmpField_Mask And &H1) = 0 Then
            Exit For
        End If
        tmpField_Mask = Fix(tmpField_Mask / 2)
        ShiftBits = ShiftBits + 1
    Next num
    Data = TmpData.ShiftLeft(ShiftBits)

    Address = GetSPMISrcAddress(CLng(Addr))
    Data = GetSPMISrcData(Data)
    BitField = GetSPMIBF(FieldMsk)

    With TheHdw.Protocol.ports(PortName).NWire
        .Frames("SPMI_WRBF").Fields("BF").Value = BitField
        .Frames("SPMI_WRBF").Fields("Addr").Value = Address
        .Frames("SPMI_WRBF").Fields("Data").Value = Data
        .Frames("SPMI_WRBF").Execute
    End With

    '''    TheHdw.Protocol.Ports(PortName).IdleWait (tlTriState_tlUseDefault)

    Exit Sub

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Sub Else Resume Next

End Sub

'Public Sub SPMI_PA_WRITE(Addr As Long, TmpData As SiteLong)
'___20200313, New AHB Method
Public Sub SPMI_PA_WRITE(Address_In As Variant, TmpData As SiteLong)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SPMI_PA_WRITE"

    Dim PortName  As String
    PortName = "NWIRE_SPMI"

    Dim Address   As Long
    Dim Data      As New SiteLong

    '_______________________________________________________________________
    '20200313, for new AHB method
    Dim Addr As Long
    Dim FieldMsk As Long
    If TypeName(Address_In) = "String" Then
        Call GetAHB_Add_BF_Value(CStr(Address_In), Addr, FieldMsk)
    Else
        Addr = CLng(Address_In)
    End If
    '_______________________________________________________________________

    Address = GetSPMISrcAddress(Addr)
    Data = GetSPMISrcData(TmpData)

    With TheHdw.Protocol.ports(PortName).NWire
        .Frames("SPMI_WRITE").Fields("Addr").Value = Address
        .Frames("SPMI_WRITE").Fields("Data").Value = Data
        .Frames("SPMI_WRITE").Execute
    End With

    '''    TheHdw.Protocol.Ports(PortName).IdleWait (tlTriState_tlUseDefault)

    Exit Sub

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Sub Else Resume Next

End Sub


''''''--------------- below is combine readback and change specificy bit to write whole address ''----------------------------------
'''Public Function SPMI_PA_WRITE(Addr As Long, TmpData As Variant, Optional FieldMsk As Long = 0, Optional ForceWholeWrite As Boolean = False)
'''
'''    Dim Address As Long
'''    Dim Data As New SiteLong
'''
'''    Dim PortName As String
'''    PortName = "NWIRE_SPMI"
'''
'''    Data = TmpData
'''
'''    ''  this function no Bit Field.
'''    If TheExec.TesterMode = testModeOffline Then Exit Function
'''
'''
'''    If FieldMsk = 0 Then
'''        Data = GetSPMISrcData(Data)
'''    Else
'''        Dim outData As New SiteLong
'''        Dim RealConcern_InData As New SiteLong
'''        Dim Except_ConcernData_From_ReadData As New SiteLong
'''    ''''-Read back to combine------------------------------------------------------------------------
'''        If ForceWholeWrite = False Then
'''            '' Normal setting with Bitfield
'''            SPMI_PA_READ Addr, outData
'''            Except_ConcernData_From_ReadData = outData.BitwiseAnd(FieldMsk)
'''
'''            ShiftBits = 0
'''            tmpField_Mask = FieldMsk
'''            For num = 1 To g_iAHB_BW
'''                If (tmpField_Mask And &H1) = 0 Then
'''                    Exit For
'''                End If
'''                tmpField_Mask = Fix(tmpField_Mask / 2)
'''                ShiftBits = ShiftBits + 1
'''            Next num
'''            RealConcern_InData = Data.ShiftLeft(ShiftBits).BitwiseAnd((2 ^ 8 - 1) - FieldMsk)
'''            Data = Except_ConcernData_From_ReadData.Add(RealConcern_InData)
'''            Data = GetSPMISrcData(Data)
'''        Else
'''            '' Special setting with Bitfield
'''            SPMI_PA_READ Addr, outData
'''            Except_ConcernData_From_ReadData = outData.BitwiseAnd(FieldMsk)
'''            ShiftBits = 0
'''            RealConcern_InData = Data.ShiftLeft(ShiftBits).BitwiseAnd((2 ^ 8 - 1) - FieldMsk)
'''            Data = Except_ConcernData_From_ReadData.Add(RealConcern_InData)
'''            Data = GetSPMISrcData(Data)
'''        End If
'''    End If
'''
'''
'''    Address = GetSPMISrcAddress(Addr)
'''
'''
'''    With TheHdw.Protocol.ports(PortName).NWire
'''        .Frames("SPMI_WRITE").Fields("Addr").Value = Address
'''        .Frames("SPMI_WRITE").Fields("Data").Value = Data
'''        .Frames("SPMI_WRITE").Execute
'''    End With
'''
'''    TheHdw.Protocol.ports(PortName).IdleWait
''''''    TheHdw.Protocol.Ports(PortName).IdleWait (tlTriState_tlUseDefault)
'''
'''End Function



Public Sub SPMI_PA_READ_TO_HRAM(Addr As Long, Data As SiteLong, Optional Field_Mask As Long = 0)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SPMI_PA_READ_TO_HRAM"

    Dim PortName  As String
    PortName = "NWIRE_SPMI"

    Dim Address   As Long
    Dim nWireReadIndexes As INWireHRAMTransactionIndexes
    Dim iSite     As Variant
    Dim tmpAHBData As New PinListData

    Address = GetSPMISrcAddress(Addr)

    With TheHdw.Protocol.ports(PortName).NWire
        .Frames("SPMI_READ").Fields("Addr").Value = Address
        .Frames("SPMI_READ").Execute
    End With

    tmpAHBData = TheHdw.Protocol.ports(PortName).NWire.HRAM.Transactions.Read()
    For Each iSite In TheExec.Sites.Selected
        Set nWireReadIndexes = tmpAHBData.Pins("NWIRE_SPMI").Value
        Data = nWireReadIndexes(0).Fields("Data").Value
    Next iSite

    '''    TheHdw.Protocol.Ports(PortName).IdleWait (tlTriState_tlUseDefault)

    '******************************************************************************************
    Dim BitAND As Long, Offset As Long
    Dim BitANDStr As String, CalcData As New SiteLong

    If Field_Mask > 0 Then
        BitAND = (&HFF) Xor Field_Mask
        BitANDStr = ConvertFormat_Dec2Bin_Complement(BitAND, 8)
        Offset = InStr(StrReverse(BitANDStr), "1") - 1
        CalcData = Data.BitWiseAnd(BitAND).ShiftRight(Offset)
        Data = CalcData
    End If


    For Each Site In TheExec.Sites
        TheExec.Datalog.WriteComment "Adress-h'" & Hex(Addr) & "(d'" & (Addr And &HFF) & ")/" & "Data-" & Hex(Data(Site))
        Debug.Print "Adress-h'" & Hex(Addr) & "(d'" & (Addr And &HFF) & ")/" & "Data-" & Hex(Data(Site))
    Next Site

    Exit Sub

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Sub Else Resume Next

End Sub

'Public Sub SPMI_PA_READ_TO_CRAM(Addr As Long, Data As SiteLong, Optional Field_Mask As Long = 0, Optional bDBGlog As Boolean = True)
'___20200313, New AHB Method
Public Sub SPMI_PA_READ_TO_CRAM(Address_In As Variant, Data As SiteLong, Optional Field_Mask_In As Variant = 0, Optional bDBGlog As Boolean = True)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SPMI_PA_READ_TO_CRAM"

    Dim PortName  As String
    PortName = "NWIRE_SPMI"

    Dim Address   As Long
    Dim nWireReadIndexes As INWireCMEMTransactionIndexes
    Dim iSite     As Variant
    Dim tmpAHBData As New PinListData
    
    '_______________________________________________________________________
    '20200313, for new AHB method
    Dim Addr As Long
    Dim Field_Mask As Long
    If TypeName(Address_In) = "String" Then
        Call GetAHB_Add_BF_Value(CStr(Address_In), Addr, Field_Mask)
    Else
        Addr = CLng(Address_In)
        Field_Mask = CLng(Field_Mask_In)
    End If
    '_______________________________________________________________________

    Address = GetSPMISrcAddress(Addr)

    If TheExec.TesterMode = testModeOffline Then Exit Sub   '20190909 avoid offline error

    With TheHdw.Protocol.ports(PortName).NWire
        .Frames("SPMI_READ").Fields("Addr").Value = Address
        .Frames("SPMI_READ").Execute tlNWireExecutionType_CaptureInCMEM
    End With

    tmpAHBData = TheHdw.Protocol.ports(PortName).NWire.CMEM.Transactions.Read()
    For Each iSite In TheExec.Sites.Selected
        Set nWireReadIndexes = tmpAHBData.Pins("NWIRE_SPMI").Value
        Data = nWireReadIndexes(0).Fields("Data").Value
    Next iSite

    '''    TheHdw.Protocol.Ports(PortName).IdleWait (tlTriState_tlUseDefault)



    '******************************************************************************************
    Dim BitAND As Long, Offset As Long
    Dim BitANDStr As String, CalcData As New SiteLong

    If Field_Mask > 0 Then
        BitAND = (&HFF) Xor Field_Mask
        BitANDStr = ConvertFormat_Dec2Bin_Complement(BitAND, 8)
        Offset = InStr(StrReverse(BitANDStr), "1") - 1
        CalcData = Data.BitWiseAnd(BitAND).ShiftRight(Offset)
        Data = CalcData
    End If



    For Each Site In TheExec.Sites
        If bDBGlog = True Then TheExec.Datalog.WriteComment "Adress-h'" & Hex(Addr) & "(d'" & (Addr And &HFF) & ")/" & "Data-" & Hex(Data(Site))
        If bDBGlog = True Then Debug.Print "Adress-h'" & Hex(Addr) & "(d'" & (Addr And &HFF) & ")/" & "Data-" & Hex(Data(Site))
    Next Site

    Exit Sub

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Sub Else Resume Next

End Sub

Public Function SPMI_WRITE(Address As Long, ByVal TmpData As SiteLong, Optional Field_Mask As Long = 0)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SPMI_WRITE"
    
    Dim Data      As New SiteLong

    Data = GetSPMISrcData(TmpData)
    Call Data_write_spmi(Address, Data)

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function SPMI_READ(Address As Long, Data As SiteLong, Optional Field_Mask As Long = 0)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SPMI_READ"
    
    Call Data_read_spmi(Address, Data)

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function GetSPMIBF(FieldMsk As Long) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetSPMIBF"
    
    Dim lParity   As Long
    Dim Index     As Integer

    lParity = cParity
    'only calc which bit is 1 per 8 bits and got final parity for data[7:0]
    For Index = 0 To 7 Step 1
        If (FieldMsk And 2 ^ Index) >= 1 Then
            lParity = (Not lParity) And BIT0
        End If
    Next Index

    GetSPMIBF = (FieldMsk * 2) + lParity

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function GetSPMISrcData(Data As SiteLong) As SiteLong

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetSPMISrcData"
    
    Dim iSite     As Variant
    Dim lParity   As Long
    Dim Index     As Integer
    Dim slRtnData As New SiteLong

    lParity = cParity

    For Each iSite In TheExec.Sites.Selected
        'only calc which bit is 1 per 8 bits and got final parity for data[7:0]
        For Index = 0 To 7 Step 1
            If (Data And 2 ^ Index) >= 1 Then
                lParity = (Not lParity) And BIT0
            End If
        Next Index

        slRtnData(iSite) = Data.Multiply(2).Add(lParity)(iSite)
        lParity = cParity
    Next iSite

    Set GetSPMISrcData = slRtnData
    
    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function GetSPMISrcAddress(Address As Long)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GetSPMISrcAddress"
    
    Dim lLow8bis As Long, lHigh8Bis As Long
    Dim lLowParity As Long, lHighParity As Long
    Dim lLow8bitsMsk As Long: lLow8bitsMsk = 255
    Dim lHigh8bitsMsk As Long: lHigh8bitsMsk = 65280

    Dim Index     As Integer

    lLow8bis = Address And lLow8bitsMsk
    lHigh8Bis = ((Address And lHigh8bitsMsk) / (lLow8bitsMsk + 1))

    lLowParity = cParity
    lHighParity = cParity

    'only calc which bit is 1 per 8 bits and got final parity for address[15:8] and address[7:0]
    For Index = 0 To 7 Step 1
        If (lLow8bis And 2 ^ Index) >= 1 Then
            lLowParity = (Not lLowParity) And BIT0
        End If

        If (lHigh8Bis And 2 ^ Index) >= 1 Then
            lHighParity = (Not lHighParity) And BIT0
        End If
    Next Index

    'address[15:8] have to shift right 2 bits for both high and low parity and address[7:0] shift lefe 1 bit for low parity
    GetSPMISrcAddress = (lHigh8Bis * (lLow8bitsMsk + 1) * 4) + lHighParity * 2 ^ 9 + lLow8bis * 2 + lLowParity

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Data_read_spmi(lAddress As Long, ByRef Data As SiteLong)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Data_read_spmi"
    
    Dim addrwidth As Long
    Dim DataWidth As Long
    Dim SPMI_PIN  As String
    Dim PatName   As New PatternSet

    Dim i         As Long
    Dim dataout   As New SiteLong
    Dim Address   As New SiteLong

    If TheExec.Sites.Active.Count = 0 Then Exit Function

    SPMI_PIN = "SPMI_SDATA"
    PatName.Value = ".\SPMI_PAT\spmi_read.pat"

    Address = GetSPMISrcAddress(lAddress)
    '''    Address = ToSiteLong(Addr)

    addrwidth = 2
    DataWidth = 1

    Dim addressSerial As New DSPWave, AdressWave As New DSPWave
    Dim arrAddr(1) As Long
    Dim lLow9bitsMsk As Long: lLow9bitsMsk = 511
    Dim lHigh9bitsMsk As Long: lHigh9bitsMsk = 130816

    For Each Site In TheExec.Sites
        addressSerial.CreateConstant Address, 2, DspDouble
        arrAddr(0) = ((Address And lHigh9bitsMsk) / (lLow9bitsMsk + 1))
        arrAddr(1) = Address And lLow9bitsMsk
        addressSerial.Data = arrAddr
        AdressWave = addressSerial.Copy
    Next Site

    Dim SignalName As String
    Dim WaveDef   As String
    WaveDef = "Wavedef"
    SignalName = "OnlyAdress"

    addrwidth = 2

    TheHdw.Patterns(PatName).Load

    For Each Site In TheExec.Sites.Selected
        TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & Site, AdressWave, True
        TheHdw.DSSC.Pins(SPMI_PIN).Pattern(PatName).Source.Signals.Add SignalName
        With TheHdw.DSSC(SPMI_PIN).Pattern(PatName).Source.Signals(SignalName)
            .Reinitialize
            .WaveDefinitionName = WaveDef & Site
            .Amplitude = 1
            .SampleSize = addrwidth    'dressWave.SampleSize 'addrwidth '20171118
            '.LoadSamples  'B0 TTR 2019-01-09
            .LoadSettings
        End With
    Next Site

    TheHdw.DSSC(SPMI_PIN).Pattern(PatName).Source.Signals.DefaultSignal = SignalName

    'setup capture
    Dim DataArray As New DSPWave
    DataWidth = 1
    DataArray.CreateConstant 0, 1, DspLong
    Call DSSC_Capture_Setup(PatName, SPMI_PIN, "dataSigAHB", DataWidth, DataArray)

    ' Bypass DSP computing, use HOST computer
    'TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
    ' Halt on opcode to make sure all samples are capture.
    TheHdw.Digital.Patgen.HaltWait

    For Each Site In TheExec.Sites
        If DataArray.SampleSize <> 8 Then
            TheExec.Datalog.WriteComment "DataArray.SampleSize=" & CStr(DataArray.SampleSize)    '20171118
            'Stop
        End If
        Data(Site) = DataArray.Element(0)
    Next Site

    For Each Site In TheExec.Sites
        TheExec.Datalog.WriteComment "Adress-h'" & Hex(lAddress) & "(d'" & (lAddress And &HFF) & ")/" & "Data-" & Hex(Data(Site))
        Debug.Print "Adress-h'" & Hex(lAddress) & "(d'" & (lAddress And &HFF) & ")/" & "Data-" & Hex(Data(Site))
    Next Site

    Exit Function


ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function

Public Function Data_write_spmi(lAddress As Long, Data As SiteLong) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Data_write_spmi"
    
    Dim addrSerial As New DSPWave
    Dim dataSerial As New DSPWave
    Dim AddrWData As New DSPWave
    Dim arrAddrWdata(2) As Long
    Dim addressplusdataWave As New DSPWave
    Dim SignalName As String
    Dim WaveDef   As String
    Dim addrwidth As Long
    Dim DataWidth As Long
    Dim i         As Long
    Dim dataout   As New SiteLong
    Dim Address   As New SiteLong
    Dim Addr      As Long

    Dim lLow9bitsMsk As Long: lLow9bitsMsk = 511
    Dim lHigh9bitsMsk As Long: lHigh9bitsMsk = 130816

    Dim TrimPattern As New PatternSet
    Dim SPMI_PIN  As String

    TrimPattern.Value = ".\SPMI_PAT\spmi_write.pat"
    SPMI_PIN = "SPMI_SDATA"

    addrwidth = 2
    DataWidth = 1

    Addr = GetSPMISrcAddress(lAddress)

    For Each Site In TheExec.Sites
        AddrWData.CreateConstant 0, 3, DspLong
    Next Site

    For Each Site In TheExec.Sites
        arrAddrWdata(0) = ((Addr And lHigh9bitsMsk) / (lLow9bitsMsk + 1))
        arrAddrWdata(1) = Addr And lLow9bitsMsk
        arrAddrWdata(2) = Data
        AddrWData.Data = arrAddrWdata
    Next Site

    WaveDef = "WaveDef"
    SignalName = "Addressplusdata"

    TheHdw.Patterns(TrimPattern).Load
    For Each Site In TheExec.Sites
        TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & Site, AddrWData, True
        TheHdw.DSSC.Pins(SPMI_PIN).Pattern(TrimPattern).Source.Signals.Add SignalName
        With TheHdw.DSSC.Pins(SPMI_PIN).Pattern(TrimPattern).Source.Signals(SignalName)
            .WaveDefinitionName = WaveDef & Site
            .SampleSize = (addrwidth + DataWidth)
            .Amplitude = 1
            '.LoadSamples  'B0 TTR 2019-01-09
            .LoadSettings
        End With
    Next Site

    TheHdw.DSSC.Pins(SPMI_PIN).Pattern(TrimPattern).Source.Signals.DefaultSignal = SignalName

    TheHdw.Patterns(TrimPattern).Start ("")

    ' Bypass DSP computing, use HOST computer
    'TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
    ' Halt on opcode to make sure all samples are capture.
    TheHdw.Digital.Patgen.HaltWait

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function

