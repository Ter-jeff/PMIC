Attribute VB_Name = "VDD_Level_From_TestPlan"
Public Type T_VDDpin
    Name As String
    VOL As Double
End Type

Public Type T_VDDpin_All
    SEQ1_VDD_MAIN_SNS_UVI80_pin As T_VDDpin
    SEQ1_VDD_MAIN_SNS_WLED_UVI80_pin As T_VDDpin
    SEQ1_VDD_MAIN_UVI80_pin As T_VDDpin
    SEQ1_VDD_MAIN_WLED_UVI80_pin As T_VDDpin
    SEQ1_VDD_MAIN_DRV_UVI80_pin As T_VDDpin
    SEQ1_VDD_MAIN_BSTLQ_UVI80_pin As T_VDDpin
    SEQ1_VDD_MAIN_LDO_UVI80_pin As T_VDDpin
    SEQ1_VDD_BOOST_UVI80_pin As T_VDDpin
    SEQ1_VDD_BOOST1_UVI80_pin As T_VDDpin
    SEQ1_VDD_BOOST_LDO_UVI80_pin As T_VDDpin
    SEQ1_VDD_BOOST_LNV_UVI80_pin As T_VDDpin
    SEQ1_VDD_BUCK2_UVI80_pin As T_VDDpin
    SEQ1_VDD_BUCK6_UVI80_pin As T_VDDpin
    SEQ1_VCP_UVI80_pin As T_VDDpin
    SEQ1_Temp _UVI80_pin As T_VDDpin
    SEQ2_VDD_ANA_UVI80_pin As T_VDDpin
    SEQ2_VDD_DIG_UVI80_pin As T_VDDpin
    SEQ2_VDD_B4_UVI80_pin As T_VDDpin
    SEQ2_VPP_UVI80_pin As T_VDDpin
    SEQ2_VDDIO1V2_UVI80_pin As T_VDDpin
    SEQ2_CRASH_L_UVI80_pin As T_VDDpin
    SEQ2_BUCK_SWI0_IN_UVI80_pin As T_VDDpin
    SEQ2_BUCK_SWI1_IN_UVI80_pin As T_VDDpin
    SEQ3_VDD_LDO14_UVI80_pin As T_VDDpin
    SEQ3_VDD_B12_LDO_UVI80_pin As T_VDDpin
    SEQ3_VDD_B6_LDO_UVI80_pin As T_VDDpin
    SEQ3_VDD_LDO6_UVI80_pin As T_VDDpin
    SEQ3_VDD_LDO7_UVI80_pin As T_VDDpin
    SEQ3_VREF_ADC_DC30_pin As T_VDDpin
End Type

Public VDDpin_All As T_VDDpin_All

Public Function ACORE_PowerUp()

    On Error GoTo ErrHandler
    Dim MeasPin As String
    Dim PinName As String
    Dim SEQ_PinName_All                         As String
    Dim RampStep                                As Integer
    Dim RampStepSize                            As Integer

    RampStepSize = 10

    '-----------Relay Connect
    '--------------------------------------------------------------------
    ' Power-up sequence 0: Connect/disconnect instrument,relay
    '--------------------------------------------------------------------
    'By diff. project setting
    'TheHdw.PPMU.Pins("ALL_DIG_PINS_NO_FRC").Disconnect

    '--------------------------------------------------------------------
    ' Power levels definition
    '--------------------------------------------------------------------
        If theexec.DataManager.InstanceName Like "*NV*" Then
            VDDpin_All.SEQ1_VDD_MAIN_SNS_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_MAIN_SNS_WLED_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_MAIN_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_MAIN_WLED_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_MAIN_DRV_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_MAIN_BSTLQ_UVI80_pin.VOL =  2#
            VDDpin_All.SEQ1_VDD_MAIN_LDO_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_BOOST_UVI80_pin.VOL =  3#
            VDDpin_All.SEQ1_VDD_BOOST1_UVI80_pin.VOL =  3#
            VDDpin_All.SEQ1_VDD_BOOST_LDO_UVI80_pin.VOL =  3#
            VDDpin_All.SEQ1_VDD_BOOST_LNV_UVI80_pin.VOL =  3#
            VDDpin_All.SEQ1_VDD_BUCK2_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_BUCK6_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VCP_UVI80_pin.VOL = 4.5
            VDDpin_All.SEQ1_Temp _UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ2_VDD_ANA_UVI80_pin.VOL = 1.35
            VDDpin_All.SEQ2_VDD_DIG_UVI80_pin.VOL = 1.35
            VDDpin_All.SEQ2_VDD_B4_UVI80_pin.VOL =  1#
            VDDpin_All.SEQ2_VPP_UVI80_pin.VOL =  0#
            VDDpin_All.SEQ2_VDDIO1V2_UVI80_pin.VOL = 1.2
            VDDpin_All.SEQ2_CRASH_L_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ2_BUCK_SWI0_IN_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ2_BUCK_SWI1_IN_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ3_VDD_LDO14_UVI80_pin.VOL = 1.15
            VDDpin_All.SEQ3_VDD_B12_LDO_UVI80_pin.VOL = 0.8
            VDDpin_All.SEQ3_VDD_B6_LDO_UVI80_pin.VOL = 1.15
            VDDpin_All.SEQ3_VDD_LDO6_UVI80_pin.VOL = 1.15
            VDDpin_All.SEQ3_VDD_LDO7_UVI80_pin.VOL =  2.5  '(TBD)
            VDDpin_All.SEQ3_VREF_ADC_DC30_pin.VOL = 1.5
            g_Voltage_Corner = "NV"
        ElseIf theexec.DataManager.InstanceName Like "*LV*" Then
            VDDpin_All.SEQ1_VDD_MAIN_SNS_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_MAIN_SNS_WLED_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_MAIN_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_MAIN_WLED_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_MAIN_DRV_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_MAIN_BSTLQ_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_MAIN_LDO_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_BOOST_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_BOOST1_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_BOOST_LDO_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_BOOST_LNV_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_BUCK2_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VDD_BUCK6_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ1_VCP_UVI80_pin.VOL =  5#
            VDDpin_All.SEQ1_Temp _UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ2_VDD_ANA_UVI80_pin.VOL = 1.5
            VDDpin_All.SEQ2_VDD_DIG_UVI80_pin.VOL = 1.5
            VDDpin_All.SEQ2_VDD_B4_UVI80_pin.VOL = 1.1
            VDDpin_All.SEQ2_VPP_UVI80_pin.VOL =  0#
            VDDpin_All.SEQ2_VDDIO1V2_UVI80_pin.VOL = 1.2
            VDDpin_All.SEQ2_CRASH_L_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ2_BUCK_SWI0_IN_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ2_BUCK_SWI1_IN_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ3_VDD_LDO14_UVI80_pin.VOL = 1.2
            VDDpin_All.SEQ3_VDD_B12_LDO_UVI80_pin.VOL = 0.9
            VDDpin_All.SEQ3_VDD_B6_LDO_UVI80_pin.VOL = 1.2
            VDDpin_All.SEQ3_VDD_LDO6_UVI80_pin.VOL = 1.2
            VDDpin_All.SEQ3_VDD_LDO7_UVI80_pin.VOL = 3.8
            VDDpin_All.SEQ3_VREF_ADC_DC30_pin.VOL = 1.5
            g_Voltage_Corner = "LV"
        ElseIf theexec.DataManager.InstanceName Like "*HV*" Then
            VDDpin_All.SEQ1_VDD_MAIN_SNS_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_MAIN_SNS_WLED_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_MAIN_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_MAIN_WLED_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_MAIN_DRV_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_MAIN_BSTLQ_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_MAIN_LDO_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_BOOST_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_BOOST1_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_BOOST_LDO_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_BOOST_LNV_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_BUCK2_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VDD_BUCK6_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ1_VCP_UVI80_pin.VOL =  5#
            VDDpin_All.SEQ1_Temp _UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ2_VDD_ANA_UVI80_pin.VOL = 1.65
            VDDpin_All.SEQ2_VDD_DIG_UVI80_pin.VOL = 1.65
            VDDpin_All.SEQ2_VDD_B4_UVI80_pin.VOL = 1.4
            VDDpin_All.SEQ2_VPP_UVI80_pin.VOL =  0#
            VDDpin_All.SEQ2_VDDIO1V2_UVI80_pin.VOL = 1.2
            VDDpin_All.SEQ2_CRASH_L_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ2_BUCK_SWI0_IN_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ2_BUCK_SWI1_IN_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ3_VDD_LDO14_UVI80_pin.VOL = 1.4
            VDDpin_All.SEQ3_VDD_B12_LDO_UVI80_pin.VOL = 1.4
            VDDpin_All.SEQ3_VDD_B6_LDO_UVI80_pin.VOL = 1.4
            VDDpin_All.SEQ3_VDD_LDO6_UVI80_pin.VOL = 1.4
            VDDpin_All.SEQ3_VDD_LDO7_UVI80_pin.VOL = 4.8
            VDDpin_All.SEQ3_VREF_ADC_DC30_pin.VOL = 1.5
            g_Voltage_Corner = "HV"
        Else    'if no any g_Voltage_Corner setting or setting error will set level into NV
            VDDpin_All.SEQ1_VDD_MAIN_SNS_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_MAIN_SNS_WLED_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_MAIN_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_MAIN_WLED_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_MAIN_DRV_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_MAIN_BSTLQ_UVI80_pin.VOL =  2#
            VDDpin_All.SEQ1_VDD_MAIN_LDO_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_BOOST_UVI80_pin.VOL =  3#
            VDDpin_All.SEQ1_VDD_BOOST1_UVI80_pin.VOL =  3#
            VDDpin_All.SEQ1_VDD_BOOST_LDO_UVI80_pin.VOL =  3#
            VDDpin_All.SEQ1_VDD_BOOST_LNV_UVI80_pin.VOL =  3#
            VDDpin_All.SEQ1_VDD_BUCK2_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VDD_BUCK6_UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ1_VCP_UVI80_pin.VOL = 4.5
            VDDpin_All.SEQ1_Temp _UVI80_pin.VOL = 2.5
            VDDpin_All.SEQ2_VDD_ANA_UVI80_pin.VOL = 1.35
            VDDpin_All.SEQ2_VDD_DIG_UVI80_pin.VOL = 1.35
            VDDpin_All.SEQ2_VDD_B4_UVI80_pin.VOL =  1#
            VDDpin_All.SEQ2_VPP_UVI80_pin.VOL =  0#
            VDDpin_All.SEQ2_VDDIO1V2_UVI80_pin.VOL = 1.2
            VDDpin_All.SEQ2_CRASH_L_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ2_BUCK_SWI0_IN_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ2_BUCK_SWI1_IN_UVI80_pin.VOL = 1.8
            VDDpin_All.SEQ3_VDD_LDO14_UVI80_pin.VOL = 1.15
            VDDpin_All.SEQ3_VDD_B12_LDO_UVI80_pin.VOL = 0.8
            VDDpin_All.SEQ3_VDD_B6_LDO_UVI80_pin.VOL = 1.15
            VDDpin_All.SEQ3_VDD_LDO6_UVI80_pin.VOL = 1.15
            VDDpin_All.SEQ3_VDD_LDO7_UVI80_pin.VOL =  2.5  '(TBD)
            VDDpin_All.SEQ3_VREF_ADC_DC30_pin.VOL = 1.5
            g_Voltage_Corner = "NV"
        End If

    '--------------------------------------------------------------------
    ' Power-up sequence 1: Apply power
    '--------------------------------------------------------------------

    '--------------------------------------------------------------------
    ' Define each pin name
    '--------------------------------------------------------------------
    VDDpin_All.SEQ1_VDD_MAIN_SNS_UVI80_pin.Name = "VDD_MAIN_SNS_UVI80"
    VDDpin_All.SEQ1_VDD_MAIN_SNS_WLED_UVI80_pin.Name = "VDD_MAIN_SNS_WLED_UVI80"
    VDDpin_All.SEQ1_VDD_MAIN_UVI80_pin.Name = "VDD_MAIN_UVI80"
    VDDpin_All.SEQ1_VDD_MAIN_WLED_UVI80_pin.Name = "VDD_MAIN_WLED_UVI80"
    VDDpin_All.SEQ1_VDD_MAIN_DRV_UVI80_pin.Name = "VDD_MAIN_DRV_UVI80"
    VDDpin_All.SEQ1_VDD_MAIN_BSTLQ_UVI80_pin.Name = "VDD_MAIN_BSTLQ_UVI80"
    VDDpin_All.SEQ1_VDD_MAIN_LDO_UVI80_pin.Name = "VDD_MAIN_LDO_UVI80"
    VDDpin_All.SEQ1_VDD_BOOST_UVI80_pin.Name = "VDD_BOOST_UVI80"
    VDDpin_All.SEQ1_VDD_BOOST1_UVI80_pin.Name = "VDD_BOOST1_UVI80"
    VDDpin_All.SEQ1_VDD_BOOST_LDO_UVI80_pin.Name = "VDD_BOOST_LDO_UVI80"
    VDDpin_All.SEQ1_VDD_BOOST_LNV_UVI80_pin.Name = "VDD_BOOST_LNV_UVI80"
    VDDpin_All.SEQ1_VDD_BUCK2_UVI80_pin.Name = "VDD_BUCK2_UVI80"
    VDDpin_All.SEQ1_VDD_BUCK6_UVI80_pin.Name = "VDD_BUCK6_UVI80"
    VDDpin_All.SEQ1_VCP_UVI80_pin.Name = "VCP_UVI80"
    VDDpin_All.SEQ1_Temp _UVI80_pin.Name = "Temp _UVI80"
    VDDpin_All.SEQ2_VDD_ANA_UVI80_pin.Name = "VDD_ANA_UVI80"
    VDDpin_All.SEQ2_VDD_DIG_UVI80_pin.Name = "VDD_DIG_UVI80"
    VDDpin_All.SEQ2_VDD_B4_UVI80_pin.Name = "VDD_B4_UVI80"
    VDDpin_All.SEQ2_VPP_UVI80_pin.Name = "VPP_UVI80"
    VDDpin_All.SEQ2_VDDIO1V2_UVI80_pin.Name = "VDDIO1V2_UVI80"
    VDDpin_All.SEQ2_CRASH_L_UVI80_pin.Name = "CRASH_L_UVI80"
    VDDpin_All.SEQ2_BUCK_SWI0_IN_UVI80_pin.Name = "BUCK_SWI0_IN_UVI80"
    VDDpin_All.SEQ2_BUCK_SWI1_IN_UVI80_pin.Name = "BUCK_SWI1_IN_UVI80"
    VDDpin_All.SEQ3_VDD_LDO14_UVI80_pin.Name = "VDD_LDO14_UVI80"
    VDDpin_All.SEQ3_VDD_B12_LDO_UVI80_pin.Name = "VDD_B12_LDO_UVI80"
    VDDpin_All.SEQ3_VDD_B6_LDO_UVI80_pin.Name = "VDD_B6_LDO_UVI80"
    VDDpin_All.SEQ3_VDD_LDO6_UVI80_pin.Name = "VDD_LDO6_UVI80"
    VDDpin_All.SEQ3_VDD_LDO7_UVI80_pin.Name = "VDD_LDO7_UVI80"
    VDDpin_All.SEQ3_VREF_ADC_DC30_pin.Name = "VREF_ADC_DC30"

    '--------------------------------------------------------------------
    ' Set all pin be a group and init pin setting
    '--------------------------------------------------------------------
    Dim MyArray(28) As String
    MyArray(0) =VDDpin_All.SEQ1_VDD_MAIN_SNS_UVI80_pin.Name
    MyArray(1) =VDDpin_All.SEQ1_VDD_MAIN_SNS_WLED_UVI80_pin.Name
    MyArray(2) =VDDpin_All.SEQ1_VDD_MAIN_UVI80_pin.Name
    MyArray(3) =VDDpin_All.SEQ1_VDD_MAIN_WLED_UVI80_pin.Name
    MyArray(4) =VDDpin_All.SEQ1_VDD_MAIN_DRV_UVI80_pin.Name
    MyArray(5) =VDDpin_All.SEQ1_VDD_MAIN_BSTLQ_UVI80_pin.Name
    MyArray(6) =VDDpin_All.SEQ1_VDD_MAIN_LDO_UVI80_pin.Name
    MyArray(7) =VDDpin_All.SEQ1_VDD_BOOST_UVI80_pin.Name
    MyArray(8) =VDDpin_All.SEQ1_VDD_BOOST1_UVI80_pin.Name
    MyArray(9) =VDDpin_All.SEQ1_VDD_BOOST_LDO_UVI80_pin.Name
    MyArray(10) =VDDpin_All.SEQ1_VDD_BOOST_LNV_UVI80_pin.Name
    MyArray(11) =VDDpin_All.SEQ1_VDD_BUCK2_UVI80_pin.Name
    MyArray(12) =VDDpin_All.SEQ1_VDD_BUCK6_UVI80_pin.Name
    MyArray(13) =VDDpin_All.SEQ1_VCP_UVI80_pin.Name
    MyArray(14) =VDDpin_All.SEQ1_Temp _UVI80_pin.Name
    MyArray(15) =VDDpin_All.SEQ2_VDD_ANA_UVI80_pin.Name
    MyArray(16) =VDDpin_All.SEQ2_VDD_DIG_UVI80_pin.Name
    MyArray(17) =VDDpin_All.SEQ2_VDD_B4_UVI80_pin.Name
    MyArray(18) =VDDpin_All.SEQ2_VPP_UVI80_pin.Name
    MyArray(19) =VDDpin_All.SEQ2_VDDIO1V2_UVI80_pin.Name
    MyArray(20) =VDDpin_All.SEQ2_CRASH_L_UVI80_pin.Name
    MyArray(21) =VDDpin_All.SEQ2_BUCK_SWI0_IN_UVI80_pin.Name
    MyArray(22) =VDDpin_All.SEQ2_BUCK_SWI1_IN_UVI80_pin.Name
    MyArray(23) =VDDpin_All.SEQ3_VDD_LDO14_UVI80_pin.Name
    MyArray(24) =VDDpin_All.SEQ3_VDD_B12_LDO_UVI80_pin.Name
    MyArray(25) =VDDpin_All.SEQ3_VDD_B6_LDO_UVI80_pin.Name
    MyArray(26) =VDDpin_All.SEQ3_VDD_LDO6_UVI80_pin.Name
    MyArray(27) =VDDpin_All.SEQ3_VDD_LDO7_UVI80_pin.Name
    MyArray(28) =VDDpin_All.SEQ3_VREF_ADC_DC30_pin.Name
    SEQ_PinName_All = Join(MyArray,",")

    TheHdw.DCVI.Pins(SEQ_PinName_All).Gate = False
    TheHdw.DCVI.Pins(SEQ_PinName_All).SetCurrentAndRange 100# * mA, 100 * mA 'Current range by diff. pin define, default setting was 100mA
    TheHdw.DCVI.Pins(SEQ_PinName_All).Voltage = 0#
    TheHdw.DCVI.Pins(SEQ_PinName_All).Connect
    TheHdw.DCVI.Pins(SEQ_PinName_All).Gate = True

    'Power up Seq 1/1/1/1/1/1/1/1/1/1/1/1/1/1/1/2/2/2/2/2/2/2/2/3/3/3/3/3/3:
    'special case : If there are speical pin, need to control by user.
    '==================================================================
    '   Step1:
    '==================================================================
    ' combine same sequence pin as same group
    For RampStep = 0 To RampStepSize Step 1
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_MAIN_SNS_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_MAIN_SNS_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_MAIN_SNS_WLED_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_MAIN_SNS_WLED_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_MAIN_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_MAIN_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_MAIN_WLED_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_MAIN_WLED_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_MAIN_DRV_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_MAIN_DRV_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_MAIN_BSTLQ_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_MAIN_BSTLQ_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_MAIN_LDO_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_MAIN_LDO_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_BOOST_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_BOOST_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_BOOST1_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_BOOST1_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_BOOST_LDO_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_BOOST_LDO_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_BOOST_LNV_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_BOOST_LNV_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_BUCK2_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_BUCK2_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VDD_BUCK6_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VDD_BUCK6_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_VCP_UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_VCP_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ1_Temp _UVI80_pin.Name).Voltage = VDDpin_All.SEQ1_Temp _UVI80.VOL * (RampStep / RampStepSize)
    Next RampStep
    TheHdw.Wait 3 * ms


    'special case : If there are speical pin, need to control by user.
    '==================================================================
    '   Step2:
    '==================================================================
    ' combine same sequence pin as same group
    For RampStep = 0 To RampStepSize Step 1
        TheHdw.DCVI.Pins(VDDpin_All.SEQ2_VDD_ANA_UVI80_pin.Name).Voltage = VDDpin_All.SEQ2_VDD_ANA_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ2_VDD_DIG_UVI80_pin.Name).Voltage = VDDpin_All.SEQ2_VDD_DIG_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ2_VDD_B4_UVI80_pin.Name).Voltage = VDDpin_All.SEQ2_VDD_B4_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ2_VPP_UVI80_pin.Name).Voltage = VDDpin_All.SEQ2_VPP_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ2_VDDIO1V2_UVI80_pin.Name).Voltage = VDDpin_All.SEQ2_VDDIO1V2_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ2_CRASH_L_UVI80_pin.Name).Voltage = VDDpin_All.SEQ2_CRASH_L_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ2_BUCK_SWI0_IN_UVI80_pin.Name).Voltage = VDDpin_All.SEQ2_BUCK_SWI0_IN_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ2_BUCK_SWI1_IN_UVI80_pin.Name).Voltage = VDDpin_All.SEQ2_BUCK_SWI1_IN_UVI80.VOL * (RampStep / RampStepSize)
    Next RampStep
    TheHdw.Wait 3 * ms


    'special case : If there are speical pin, need to control by user.
    '==================================================================
    '   Step3:
    '==================================================================
    ' combine same sequence pin as same group
    For RampStep = 0 To RampStepSize Step 1
        TheHdw.DCVI.Pins(VDDpin_All.SEQ3_VDD_LDO14_UVI80_pin.Name).Voltage = VDDpin_All.SEQ3_VDD_LDO14_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ3_VDD_B12_LDO_UVI80_pin.Name).Voltage = VDDpin_All.SEQ3_VDD_B12_LDO_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ3_VDD_B6_LDO_UVI80_pin.Name).Voltage = VDDpin_All.SEQ3_VDD_B6_LDO_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ3_VDD_LDO6_UVI80_pin.Name).Voltage = VDDpin_All.SEQ3_VDD_LDO6_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ3_VDD_LDO7_UVI80_pin.Name).Voltage = VDDpin_All.SEQ3_VDD_LDO7_UVI80.VOL * (RampStep / RampStepSize)
        TheHdw.DCVI.Pins(VDDpin_All.SEQ3_VREF_ADC_DC30_pin.Name).Voltage = VDDpin_All.SEQ3_VREF_ADC_DC30.VOL * (RampStep / RampStepSize)
    Next RampStep
    TheHdw.Wait 3 * ms



    '*********************** The area control FRC by device user need to move FRC control sequence *********************************
    'Call Acore.FreeRunningClockStop
    'TheHdw.Wait 5 * ms
    'TheHdw.Digital.Pins("ALL_DIG_PINS_NO_FRC").Connect
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, , "ALL_DIG_PINS_NO_FRC,VSS_DFT_2"  'VSS_DFT_2 low load OTP, high not load OTP

    '--------------------------------------------------------------------
    ' Start Free Running Clock 32KHz
    ' Connect and start XOUT clock
    '--------------------------------------------------------------------
    If UCase(TheExec.DataManager.InstanceName) Like "*DIGITAL*" Then
        'If user had digital pin and need to define here.(Non-Enable FRC)
        'TheHdw.Digital.Pins("XOUT_PA").Disconnect 
    Else
        'If user had digital pin and need to define here.(Enable FRC)
        'TheHdw.Digital.Pins("XOUT_PA").Disconnect
        'Call Acore.FreeRunningClockStart(32768000)
    End If
    '*********************** End of FRC control area *******************************************************************************

    'Display All Power:
    g_AllVDDPower = SEQ_PinName_All & ",CRASH_L,NXTAL_MEMS" & ",IBAT_UVI80,VBAT_UVI80"
   
    Exit Function

ErrHandler:
   Debug.Print Err.Description
   Stop
   Resume
End Function

