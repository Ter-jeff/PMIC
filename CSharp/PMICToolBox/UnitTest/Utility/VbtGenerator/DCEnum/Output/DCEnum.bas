Attribute VB_Name = "DCEnum"
Public Enum SP_Conti_Pins
    AMUX_AY = 0
    TCAL = 1
    TDEV1 = 2
    ADC_IN_UVI80 = 3
    BRICK_ID1_UVI80 = 4
    BRICK_ID2_UVI80 = 5
    VREF_UVI80 = 6
    WLED_HP1_LX_DC30 = 7
    WLED_HP2_LX_DC30 = 8
    WLED_LP_LX_DC30 = 9
    BSTLQ_LX_UVI80 = 10
    BSTLQ_VOUT_UVI80 = 11
    BUCK6_LX_UVI80 = 12
    IREF_UVI80 = 13
    RC_VOUT_UVI80 = 14
    VDD_BUCK2_UVI80 = 15
    VDD_BUCK6_UVI80 = 16
    VCP_UVI80 = 17
    BUCK_SWI0_IN_UVI80 = 18
    BUCK_SWI1_IN_UVI80 = 19
    VLDO1_UVI80 = 20
    VLDO2_UVI80 = 21
    VLDO3_UVI80 = 22
    VLDO4_UVI80 = 23
    VLDO5_UVI80 = 24
    VLDO6_UVI80 = 25
    VLDO7_UVI80 = 26
    VLDO8_UVI80 = 27
    VLDO9_UVI80 = 28
    VLDO10_UVI80 = 29
    VLDO11_UVI80 = 30
    VLDO12_UVI80 = 31
    VLDO13_UVI80 = 32
    VLDO14_UVI80 = 33
    VLDOINT_UVI80 = 34
    IDAC_5_0_BUS_UVI80 = 35
    IDAC_11_6_BUS_UVI80 = 36
    IDAC_17_12_BUS_UVI80 = 37
End Enum

Public Function SP_Conti_Pins_Cond(idx As Double) As String

    Dim SP_PinName, SPForceI, Wait_Time, On_Relay, Off_Relay, MustDiscnctPins, SPCondPin, SPCondPinV_I, TestItem, Output As String

    Select Case idx

    Case 0                                                                     '<Comment>AMUX_AY</Comment>
        SP_PinName = "AMUX_AY"
        SPForceI = ""                                                          '<Comment>AMUX_AY|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>AMUX_AY|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>AMUX_AY|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>AMUX_AY|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>AMUX_AY|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "AMUX_AY_DC30(DCVI)"                                       '<Comment>AMUX_AY|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "-1.5/0"                                                '<Comment>AMUX_AY|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>AMUX_AY|TestItem|TM_IIL_IIH =</Comment>

    Case 1                                                                     '<Comment>TCAL</Comment>
        SP_PinName = "TCAL"
        SPForceI = ""                                                          '<Comment>TCAL|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>TCAL|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>TCAL|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>TCAL|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>TCAL|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "TCAL_DC30(DCVI)"                                          '<Comment>TCAL|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "-1.5/0"                                                '<Comment>TCAL|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>TCAL|TestItem|TM_IIL_IIH =</Comment>

    Case 2                                                                     '<Comment>TDEV1</Comment>
        SP_PinName = "TDEV1"
        SPForceI = ""                                                          '<Comment>TDEV1|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>TDEV1|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>TDEV1|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>TDEV1|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>TDEV1|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "TDEV1_DC30(DCVI)"                                         '<Comment>TDEV1|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "-1.5/0"                                                '<Comment>TDEV1|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>TDEV1|TestItem|TM_IIL_IIH =</Comment>

    Case 3                                                                     '<Comment>ADC_IN_UVI80</Comment>
        SP_PinName = "ADC_IN_UVI80"
        SPForceI = ""                                                          '<Comment>ADC_IN_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>ADC_IN_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K6801"                                                     '<Comment>ADC_IN_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K6802,K6803"                                              '<Comment>ADC_IN_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "ADC_IN(PPMU);ADC_IN(Digital)"                       '<Comment>ADC_IN_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>ADC_IN_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>ADC_IN_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>ADC_IN_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 4                                                                     '<Comment>BRICK_ID1_UVI80</Comment>
        SP_PinName = "BRICK_ID1_UVI80"
        SPForceI = ""                                                          '<Comment>BRICK_ID1_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>BRICK_ID1_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K6802"                                                     '<Comment>BRICK_ID1_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K6801,K6803"                                              '<Comment>BRICK_ID1_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "ADC_IN(PPMU);ADC_IN(Digital)"                       '<Comment>BRICK_ID1_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>BRICK_ID1_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>BRICK_ID1_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BRICK_ID1_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 5                                                                     '<Comment>BRICK_ID2_UVI80</Comment>
        SP_PinName = "BRICK_ID2_UVI80"
        SPForceI = ""                                                          '<Comment>BRICK_ID2_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>BRICK_ID2_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K6803"                                                     '<Comment>BRICK_ID2_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K6801,K6802"                                              '<Comment>BRICK_ID2_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "ADC_IN(PPMU);ADC_IN(Digital)"                       '<Comment>BRICK_ID2_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>BRICK_ID2_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>BRICK_ID2_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BRICK_ID2_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 6                                                                     '<Comment>VREF_UVI80</Comment>
        SP_PinName = "VREF_UVI80"
        SPForceI = ""                                                          '<Comment>VREF_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>VREF_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VREF_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VREF_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "VREF_DC30(DCVI)"                                    '<Comment>VREF_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VREF_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VREF_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VREF_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 7                                                                     '<Comment>WLED_HP1_LX_DC30</Comment>
        SP_PinName = "WLED_HP1_LX_DC30"
        SPForceI = ""                                                          '<Comment>WLED_HP1_LX_DC30|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>WLED_HP1_LX_DC30|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1066,K1067"                                               '<Comment>WLED_HP1_LX_DC30|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K1064,K1065,K1068,K1069"                                  '<Comment>WLED_HP1_LX_DC30|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>WLED_HP1_LX_DC30|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>WLED_HP1_LX_DC30|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>WLED_HP1_LX_DC30|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>WLED_HP1_LX_DC30|TestItem|TM_IIL_IIH =</Comment>

    Case 8                                                                     '<Comment>WLED_HP2_LX_DC30</Comment>
        SP_PinName = "WLED_HP2_LX_DC30"
        SPForceI = ""                                                          '<Comment>WLED_HP2_LX_DC30|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>WLED_HP2_LX_DC30|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1068,K1069"                                               '<Comment>WLED_HP2_LX_DC30|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K1064,K1065,K1066,K1067"                                  '<Comment>WLED_HP2_LX_DC30|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>WLED_HP2_LX_DC30|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>WLED_HP2_LX_DC30|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>WLED_HP2_LX_DC30|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>WLED_HP2_LX_DC30|TestItem|TM_IIL_IIH =</Comment>

    Case 9                                                                     '<Comment>WLED_LP_LX_DC30</Comment>
        SP_PinName = "WLED_LP_LX_DC30"
        SPForceI = ""                                                          '<Comment>WLED_LP_LX_DC30|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>WLED_LP_LX_DC30|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1064,K1065"                                               '<Comment>WLED_LP_LX_DC30|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K1066,K1067,K1068,K1069"                                  '<Comment>WLED_LP_LX_DC30|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>WLED_LP_LX_DC30|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>WLED_LP_LX_DC30|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>WLED_LP_LX_DC30|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>WLED_LP_LX_DC30|TestItem|TM_IIL_IIH =</Comment>

    Case 10                                                                    '<Comment>BSTLQ_LX_UVI80</Comment>
        SP_PinName = "BSTLQ_LX_UVI80"
        SPForceI = ""                                                          '<Comment>BSTLQ_LX_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>BSTLQ_LX_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BSTLQ_LX_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BSTLQ_LX_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "CB_BSTLQ_LX_UVI80(Digital);CB_VDD_BSTLQ_IN_BSTLQ_VOUT(Digital)"                '<Comment>BSTLQ_LX_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "CB_BSTLQ_LX_UVI80(PPMU);CB_VDD_BSTLQ_IN_BSTLQ_VOUT(PPMU)"                            '<Comment>BSTLQ_LX_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "0;5"                                                   '<Comment>BSTLQ_LX_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BSTLQ_LX_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 11                                                                    '<Comment>BSTLQ_VOUT_UVI80</Comment>
        SP_PinName = "BSTLQ_VOUT_UVI80"
        SPForceI = ""                                                          '<Comment>BSTLQ_VOUT_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>BSTLQ_VOUT_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BSTLQ_VOUT_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BSTLQ_VOUT_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "CB_BSTLQ_LX_UVI80(Digital);CB_VDD_BSTLQ_IN_BSTLQ_VOUT(Digital)"                '<Comment>BSTLQ_VOUT_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "CB_BSTLQ_LX_UVI80(PPMU);CB_VDD_BSTLQ_IN_BSTLQ_VOUT(PPMU)"                            '<Comment>BSTLQ_VOUT_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "0;5"                                                   '<Comment>BSTLQ_VOUT_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BSTLQ_VOUT_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 12                                                                    '<Comment>BUCK6_LX_UVI80</Comment>
        SP_PinName = "BUCK6_LX_UVI80"
        SPForceI = ""                                                          '<Comment>BUCK6_LX_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>BUCK6_LX_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K3902"                                                     '<Comment>BUCK6_LX_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BUCK6_LX_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>BUCK6_LX_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>BUCK6_LX_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>BUCK6_LX_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BUCK6_LX_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 13                                                                    '<Comment>IREF_UVI80</Comment>
        SP_PinName = "IREF_UVI80"
        SPForceI = ""                                                          '<Comment>IREF_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>IREF_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K6901"                                                     '<Comment>IREF_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K6902,K6903"                                              '<Comment>IREF_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IREF_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IREF_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IREF_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IREF_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 14                                                                    '<Comment>RC_VOUT_UVI80</Comment>
        SP_PinName = "RC_VOUT_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>RC_VOUT_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>RC_VOUT_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>RC_VOUT_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>RC_VOUT_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>RC_VOUT_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>RC_VOUT_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>RC_VOUT_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>RC_VOUT_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 15                                                                    '<Comment>VDD_BUCK2_UVI80</Comment>
        SP_PinName = "VDD_BUCK2_UVI80"
        SPForceI = ""                                                          '<Comment>VDD_BUCK2_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VDD_BUCK2_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VDD_BUCK2_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VDD_BUCK2_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VDD_BUCK2_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VDD_BUCK2_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VDD_BUCK2_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VDD_BUCK2_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 16                                                                    '<Comment>VDD_BUCK6_UVI80</Comment>
        SP_PinName = "VDD_BUCK6_UVI80"
        SPForceI = ""                                                          '<Comment>VDD_BUCK6_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VDD_BUCK6_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VDD_BUCK6_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VDD_BUCK6_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VDD_BUCK6_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VDD_BUCK6_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VDD_BUCK6_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VDD_BUCK6_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 17                                                                    '<Comment>VCP_UVI80</Comment>
        SP_PinName = "VCP_UVI80"
        SPForceI = ""                                                          '<Comment>VCP_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VCP_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VCP_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VCP_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VCP_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VCP_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VCP_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VCP_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 18                                                                    '<Comment>BUCK_SWI0_IN_UVI80</Comment>
        SP_PinName = "BUCK_SWI0_IN_UVI80"
        SPForceI = ""                                                          '<Comment>BUCK_SWI0_IN_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>BUCK_SWI0_IN_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BUCK_SWI0_IN_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BUCK_SWI0_IN_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>BUCK_SWI0_IN_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>BUCK_SWI0_IN_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>BUCK_SWI0_IN_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BUCK_SWI0_IN_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 19                                                                    '<Comment>BUCK_SWI1_IN_UVI80</Comment>
        SP_PinName = "BUCK_SWI1_IN_UVI80"
        SPForceI = ""                                                          '<Comment>BUCK_SWI1_IN_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>BUCK_SWI1_IN_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BUCK_SWI1_IN_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BUCK_SWI1_IN_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>BUCK_SWI1_IN_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>BUCK_SWI1_IN_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>BUCK_SWI1_IN_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BUCK_SWI1_IN_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 20                                                                    '<Comment>VLDO1_UVI80</Comment>
        SP_PinName = "VLDO1_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO1_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO1_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO1_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO1_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO1_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO1_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO1_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO1_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 21                                                                    '<Comment>VLDO2_UVI80</Comment>
        SP_PinName = "VLDO2_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO2_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO2_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO2_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO2_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO2_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO2_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO2_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO2_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 22                                                                    '<Comment>VLDO3_UVI80</Comment>
        SP_PinName = "VLDO3_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO3_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO3_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO3_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO3_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO3_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO3_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO3_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO3_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 23                                                                    '<Comment>VLDO4_UVI80</Comment>
        SP_PinName = "VLDO4_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO4_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO4_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO4_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO4_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO4_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO4_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO4_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO4_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 24                                                                    '<Comment>VLDO5_UVI80</Comment>
        SP_PinName = "VLDO5_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO5_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO5_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO5_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO5_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO5_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO5_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO5_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO5_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 25                                                                    '<Comment>VLDO6_UVI80</Comment>
        SP_PinName = "VLDO6_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO6_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO6_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO6_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO6_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO6_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO6_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO6_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO6_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 26                                                                    '<Comment>VLDO7_UVI80</Comment>
        SP_PinName = "VLDO7_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO7_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO7_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO7_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO7_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO7_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO7_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO7_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO7_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 27                                                                    '<Comment>VLDO8_UVI80</Comment>
        SP_PinName = "VLDO8_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO8_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO8_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO8_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO8_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO8_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO8_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO8_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO8_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 28                                                                    '<Comment>VLDO9_UVI80</Comment>
        SP_PinName = "VLDO9_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO9_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO9_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO9_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO9_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO9_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO9_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO9_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO9_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 29                                                                    '<Comment>VLDO10_UVI80</Comment>
        SP_PinName = "VLDO10_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO10_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO10_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO10_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO10_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO10_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO10_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO10_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO10_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 30                                                                    '<Comment>VLDO11_UVI80</Comment>
        SP_PinName = "VLDO11_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO11_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO11_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO11_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO11_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO11_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO11_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO11_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO11_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 31                                                                    '<Comment>VLDO12_UVI80</Comment>
        SP_PinName = "VLDO12_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO12_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO12_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO12_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO12_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO12_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO12_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO12_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO12_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 32                                                                    '<Comment>VLDO13_UVI80</Comment>
        SP_PinName = "VLDO13_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO13_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO13_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO13_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO13_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO13_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO13_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO13_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO13_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 33                                                                    '<Comment>VLDO14_UVI80</Comment>
        SP_PinName = "VLDO14_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDO14_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDO14_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDO14_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDO14_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDO14_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDO14_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDO14_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDO14_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 34                                                                    '<Comment>VLDOINT_UVI80</Comment>
        SP_PinName = "VLDOINT_UVI80"
        SPForceI = "-0.01"                                                     '<Comment>VLDOINT_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = "0.005"                                                    '<Comment>VLDOINT_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VLDOINT_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VLDOINT_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VLDOINT_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VLDOINT_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VLDOINT_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VLDOINT_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 35                                                                    '<Comment>IDAC_5_0_BUS_UVI80</Comment>
        SP_PinName = "IDAC_5_0_BUS_UVI80"
        SPForceI = ""                                                          '<Comment>IDAC_5_0_BUS_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>IDAC_5_0_BUS_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "0/K1042,K1061;K1042,K1060;K1042,K1059;K1042,K1058;K1042,K1057;K1042,K1056"                                     '<Comment>IDAC_5_0_BUS_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_5_0_BUS_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_5_0_BUS_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_5_0_BUS_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_5_0_BUS_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_5_0_BUS_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 36                                                                    '<Comment>IDAC_11_6_BUS_UVI80</Comment>
        SP_PinName = "IDAC_11_6_BUS_UVI80"
        SPForceI = ""                                                          '<Comment>IDAC_11_6_BUS_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>IDAC_11_6_BUS_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "6/K1040,K1055;K1040,K1054;K1040,K1053;K1040,K1052;K1040,K1051;K1040,K1050"                                     '<Comment>IDAC_11_6_BUS_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_11_6_BUS_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_11_6_BUS_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_11_6_BUS_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_11_6_BUS_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_11_6_BUS_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 37                                                                    '<Comment>IDAC_17_12_BUS_UVI80</Comment>
        SP_PinName = "IDAC_17_12_BUS_UVI80"
        SPForceI = ""                                                          '<Comment>IDAC_17_12_BUS_UVI80|SpecificLimit|SPLimit = </Comment> 
        Wait_Time = ""                                                         '<Comment>IDAC_17_12_BUS_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "12/K1038,K1049;K1038,K1048;K1038,K1047;K1038,K1046;K1038,K1045;K1038,K1044"                                     '<Comment>IDAC_17_12_BUS_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_17_12_BUS_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_17_12_BUS_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_17_12_BUS_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_17_12_BUS_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_17_12_BUS_UVI80|TestItem|TM_IIL_IIH =</Comment>
    Case Else
             'SP_Leak_Pins_Cond = "Wrong_Enum_Input"
    End Select

    SP_Conti_Pins_Cond= SP_PinName + "&" + SPForceI + "&" + Wait_Time + "&" + On_Relay + "&" + Off_Relay + "&" + MustDiscnctPins + "&" + SPCondPin + "&" + SPCondPinV_I + "&" + TestItem

End Function


Public Function GenContiPinDic()
Dim idx As Double
Dim Pin_Num As Double
Dim PinName() As String
Dim Concat As String
Pin_Num = 38
ReDim SPPins(37) As String

Dim Dic_PinName() As String
ReDim Preserve Dic_PinName(Pin_Num)

    For idx = 0 To Pin_Num

    ReDim Preserve PinName(Pin_Num)
    PinName(idx) = Split_Concat(SP_Conti_Pins_Cond(idx), 0) '.SP_Leak_Pins

    Dic_PinName(idx) = PinName(idx)

    If ContiPinDic.Exists(Dic_PinName(idx)) Then
    Else
       ContiPinDic.Add Dic_PinName(idx), idx     'add key and item to dictionary
    End If
    Next idx

End Function


Public Function SearchDicIdx_Conti(PinName As String) As Double
Dim idx As Double
For idx = 0 To 38
    If ContiPinDic.Keys(idx) = PinName Then
    SearchDicIdx_Conti = idx
    Exit For
    End If
Next idx

End Function

Public Enum SP_Leak_Pins
    BUCK2_CBOT_UVI80 = 0
    BUCK2_CTOP_UVI80 = 1
    BUCK2_LX_UVI80 = 2
    BUCK2_FB_UVI80 = 3
    BUCK6_VOUT_UVI80 = 4
    BUCK6_LX_UVI80 = 5
    ADC_IN_UVI80 = 6
    BRICK_ID1_UVI80 = 7
    BRICK_ID2_UVI80 = 8
    IREF_UVI80 = 9
    VREF_UVI80 = 10
    WLED_VOUT_FB_DC30 = 11
    WLED_HP1_LX_DC30 = 12
    WLED_HP2_LX_DC30 = 13
    WLED_LP_LX_DC30 = 14
    BSTLQ_LX_UVI80 = 15
    BSTLQ_VOUT_UVI80 = 16
    VDD_BSTLQ_IN_UVI80 = 17
    VDD_MAIN_SNS_WLED_UVI80 = 18
    IDAC_OUT_0 = 19
    IDAC_OUT_1 = 20
    IDAC_OUT_2 = 21
    IDAC_OUT_3 = 22
    IDAC_OUT_4 = 23
    IDAC_OUT_5 = 24
    IDAC_OUT_6 = 25
    IDAC_OUT_7 = 26
    IDAC_OUT_8 = 27
    IDAC_OUT_9 = 28
    IDAC_OUT_10 = 29
    IDAC_OUT_11 = 30
    IDAC_OUT_12 = 31
    IDAC_OUT_13 = 32
    IDAC_OUT_14 = 33
    IDAC_OUT_15 = 34
    IDAC_OUT_16 = 35
    IDAC_OUT_17 = 36
End Enum

Public Function SP_Leak_Pins_Cond(idx As Double) As String

    Dim SP_PinName, SPLimit, SPIRange, Wait_Time, On_Relay, Off_Relay, MustDiscnctPins, SPCondPin, SPCondPinV_I, TestItem, Output As String

    Select Case idx

    Case 0                                                                     '<Comment>BUCK2_CBOT_UVI80</Comment>
        SP_PinName = "BUCK2_CBOT_UVI80"
        SPLimit = "0.02"                                                       '<Comment>BUCK2_CBOT_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = "0.02"                                                      '<Comment>BUCK2_CBOT_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.001"                                                    '<Comment>BUCK2_CBOT_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BUCK2_CBOT_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BUCK2_CBOT_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "BUCK2_FB_UVI80(DCVI)"                               '<Comment>BUCK2_CBOT_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "BUCK2_LX_UVI80(DCVI)"                                     '<Comment>BUCK2_CBOT_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "3.8/0.02"                                              '<Comment>BUCK2_CBOT_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "HI"                                                       '<Comment>BUCK2_CBOT_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 1                                                                     '<Comment>BUCK2_CTOP_UVI80</Comment>
        SP_PinName = "BUCK2_CTOP_UVI80"
        SPLimit = "0.02"                                                       '<Comment>BUCK2_CTOP_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = "0.02"                                                      '<Comment>BUCK2_CTOP_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.001"                                                    '<Comment>BUCK2_CTOP_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BUCK2_CTOP_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BUCK2_CTOP_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>BUCK2_CTOP_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "BUCK2_LX_UVI80(DCVI)"                                     '<Comment>BUCK2_CTOP_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "3.8/0.02"                                              '<Comment>BUCK2_CTOP_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "HI"                                                       '<Comment>BUCK2_CTOP_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 2                                                                     '<Comment>BUCK2_LX_UVI80</Comment>
        SP_PinName = "BUCK2_LX_UVI80"
        SPLimit = ""                                                           '<Comment>BUCK2_LX_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>BUCK2_LX_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.001"                                                    '<Comment>BUCK2_LX_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BUCK2_LX_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BUCK2_LX_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "BUCK2_FB_UVI80(DCVI)"                               '<Comment>BUCK2_LX_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "BUCK2_CTOP_UVI80(DCVI)"                                   '<Comment>BUCK2_LX_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "3.8/0.02"                                              '<Comment>BUCK2_LX_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "HI"                                                       '<Comment>BUCK2_LX_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 3                                                                     '<Comment>BUCK2_FB_UVI80</Comment>
        SP_PinName = "BUCK2_FB_UVI80"
        SPLimit = "0.0002"                                                     '<Comment>BUCK2_FB_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = "0.0002"                                                    '<Comment>BUCK2_FB_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = ""                                                         '<Comment>BUCK2_FB_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BUCK2_FB_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BUCK2_FB_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "BUCK2_CBOT_UVI80(DCVI)"                             '<Comment>BUCK2_FB_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>BUCK2_FB_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>BUCK2_FB_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BUCK2_FB_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 4                                                                     '<Comment>BUCK6_VOUT_UVI80</Comment>
        SP_PinName = "BUCK6_VOUT_UVI80"
        SPLimit = "0.0002"                                                     '<Comment>BUCK6_VOUT_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = "0.0002"                                                    '<Comment>BUCK6_VOUT_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.05"                                                     '<Comment>BUCK6_VOUT_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BUCK6_VOUT_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BUCK6_VOUT_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>BUCK6_VOUT_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>BUCK6_VOUT_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>BUCK6_VOUT_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BUCK6_VOUT_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 5                                                                     '<Comment>BUCK6_LX_UVI80</Comment>
        SP_PinName = "BUCK6_LX_UVI80"
        SPLimit = ""                                                           '<Comment>BUCK6_LX_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>BUCK6_LX_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = ""                                                         '<Comment>BUCK6_LX_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K3902"                                                     '<Comment>BUCK6_LX_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BUCK6_LX_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>BUCK6_LX_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>BUCK6_LX_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>BUCK6_LX_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BUCK6_LX_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 6                                                                     '<Comment>ADC_IN_UVI80</Comment>
        SP_PinName = "ADC_IN_UVI80"
        SPLimit = "0.0002"                                                     '<Comment>ADC_IN_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = "0.0002"                                                    '<Comment>ADC_IN_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = ""                                                         '<Comment>ADC_IN_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K6801"                                                     '<Comment>ADC_IN_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K6802,K6803"                                              '<Comment>ADC_IN_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "ADC_IN(PPMU);ADC_IN(Digital)"                       '<Comment>ADC_IN_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>ADC_IN_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>ADC_IN_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>ADC_IN_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 7                                                                     '<Comment>BRICK_ID1_UVI80</Comment>
        SP_PinName = "BRICK_ID1_UVI80"
        SPLimit = "0.0002"                                                     '<Comment>BRICK_ID1_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = "0.0002"                                                    '<Comment>BRICK_ID1_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = ""                                                         '<Comment>BRICK_ID1_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K6802"                                                     '<Comment>BRICK_ID1_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K6801,K6803"                                              '<Comment>BRICK_ID1_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "ADC_IN(PPMU);ADC_IN(Digital)"                       '<Comment>BRICK_ID1_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>BRICK_ID1_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>BRICK_ID1_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BRICK_ID1_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 8                                                                     '<Comment>BRICK_ID2_UVI80</Comment>
        SP_PinName = "BRICK_ID2_UVI80"
        SPLimit = "0.0002"                                                     '<Comment>BRICK_ID2_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = "0.0002"                                                    '<Comment>BRICK_ID2_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = ""                                                         '<Comment>BRICK_ID2_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K6803"                                                     '<Comment>BRICK_ID2_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K6801,K6802"                                              '<Comment>BRICK_ID2_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "ADC_IN(PPMU);ADC_IN(Digital)"                       '<Comment>BRICK_ID2_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>BRICK_ID2_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>BRICK_ID2_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BRICK_ID2_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 9                                                                     '<Comment>IREF_UVI80</Comment>
        SP_PinName = "IREF_UVI80"
        SPLimit = ""                                                           '<Comment>IREF_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IREF_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = ""                                                         '<Comment>IREF_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K6901"                                                     '<Comment>IREF_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K6902,K6903"                                              '<Comment>IREF_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IREF_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IREF_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IREF_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IREF_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 10                                                                    '<Comment>VREF_UVI80</Comment>
        SP_PinName = "VREF_UVI80"
        SPLimit = ""                                                           '<Comment>VREF_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>VREF_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = ""                                                         '<Comment>VREF_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VREF_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VREF_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VREF_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "VREF_DC30(DCVI)"                                          '<Comment>VREF_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "-1.5/0"                                                '<Comment>VREF_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VREF_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 11                                                                    '<Comment>WLED_VOUT_FB_DC30</Comment>
        SP_PinName = "WLED_VOUT_FB_DC30"
        SPLimit = ""                                                           '<Comment>WLED_VOUT_FB_DC30|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>WLED_VOUT_FB_DC30|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = ""                                                         '<Comment>WLED_VOUT_FB_DC30|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>WLED_VOUT_FB_DC30|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K1070,K1071"                                              '<Comment>WLED_VOUT_FB_DC30|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>WLED_VOUT_FB_DC30|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>WLED_VOUT_FB_DC30|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>WLED_VOUT_FB_DC30|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>WLED_VOUT_FB_DC30|TestItem|TM_IIL_IIH =</Comment>

    Case 12                                                                    '<Comment>WLED_HP1_LX_DC30</Comment>
        SP_PinName = "WLED_HP1_LX_DC30"
        SPLimit = ""                                                           '<Comment>WLED_HP1_LX_DC30|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>WLED_HP1_LX_DC30|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>WLED_HP1_LX_DC30|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1066,K1067"                                               '<Comment>WLED_HP1_LX_DC30|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K1064,K1065,K1068,K1069"                                  '<Comment>WLED_HP1_LX_DC30|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>WLED_HP1_LX_DC30|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>WLED_HP1_LX_DC30|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>WLED_HP1_LX_DC30|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>WLED_HP1_LX_DC30|TestItem|TM_IIL_IIH =</Comment>

    Case 13                                                                    '<Comment>WLED_HP2_LX_DC30</Comment>
        SP_PinName = "WLED_HP2_LX_DC30"
        SPLimit = ""                                                           '<Comment>WLED_HP2_LX_DC30|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>WLED_HP2_LX_DC30|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = ""                                                         '<Comment>WLED_HP2_LX_DC30|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1068,K1069"                                               '<Comment>WLED_HP2_LX_DC30|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K1064,K1065,K1066,K1067"                                  '<Comment>WLED_HP2_LX_DC30|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>WLED_HP2_LX_DC30|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>WLED_HP2_LX_DC30|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>WLED_HP2_LX_DC30|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>WLED_HP2_LX_DC30|TestItem|TM_IIL_IIH =</Comment>

    Case 14                                                                    '<Comment>WLED_LP_LX_DC30</Comment>
        SP_PinName = "WLED_LP_LX_DC30"
        SPLimit = ""                                                           '<Comment>WLED_LP_LX_DC30|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>WLED_LP_LX_DC30|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = ""                                                         '<Comment>WLED_LP_LX_DC30|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1064,K1065"                                               '<Comment>WLED_LP_LX_DC30|On_Relay|On_Relay =</Comment> 
        Off_Relay = "K1066,K1067,K1068,K1069"                                  '<Comment>WLED_LP_LX_DC30|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>WLED_LP_LX_DC30|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>WLED_LP_LX_DC30|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>WLED_LP_LX_DC30|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>WLED_LP_LX_DC30|TestItem|TM_IIL_IIH =</Comment>

    Case 15                                                                    '<Comment>BSTLQ_LX_UVI80</Comment>
        SP_PinName = "BSTLQ_LX_UVI80"
        SPLimit = ""                                                           '<Comment>BSTLQ_LX_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>BSTLQ_LX_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>BSTLQ_LX_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BSTLQ_LX_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BSTLQ_LX_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "CB_BSTLQ_LX_UVI80(Digital);CB_VDD_BSTLQ_IN_BSTLQ_VOUT(Digital);BSTLQ_VOUT_UVI80(DCVI)"                '<Comment>BSTLQ_LX_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "CB_BSTLQ_LX_UVI80(PPMU);CB_VDD_BSTLQ_IN_BSTLQ_VOUT(PPMU)"                            '<Comment>BSTLQ_LX_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "0;5"                                                   '<Comment>BSTLQ_LX_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BSTLQ_LX_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 16                                                                    '<Comment>BSTLQ_VOUT_UVI80</Comment>
        SP_PinName = "BSTLQ_VOUT_UVI80"
        SPLimit = ""                                                           '<Comment>BSTLQ_VOUT_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>BSTLQ_VOUT_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>BSTLQ_VOUT_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>BSTLQ_VOUT_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>BSTLQ_VOUT_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = "CB_BSTLQ_LX_UVI80(Digital);CB_VDD_BSTLQ_IN_BSTLQ_VOUT(Digital);BSTLQ_LX_UVI80(DCVI)"                '<Comment>BSTLQ_VOUT_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = "CB_BSTLQ_LX_UVI80(PPMU);CB_VDD_BSTLQ_IN_BSTLQ_VOUT(PPMU)"                            '<Comment>BSTLQ_VOUT_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = "0;5"                                                   '<Comment>BSTLQ_VOUT_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>BSTLQ_VOUT_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 17                                                                    '<Comment>VDD_BSTLQ_IN_UVI80</Comment>
        SP_PinName = "VDD_BSTLQ_IN_UVI80"
        SPLimit = ""                                                           '<Comment>VDD_BSTLQ_IN_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>VDD_BSTLQ_IN_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>VDD_BSTLQ_IN_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VDD_BSTLQ_IN_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VDD_BSTLQ_IN_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VDD_BSTLQ_IN_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VDD_BSTLQ_IN_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VDD_BSTLQ_IN_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VDD_BSTLQ_IN_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 18                                                                    '<Comment>VDD_MAIN_SNS_WLED_UVI80</Comment>
        SP_PinName = "VDD_MAIN_SNS_WLED_UVI80"
        SPLimit = ""                                                           '<Comment>VDD_MAIN_SNS_WLED_UVI80|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>VDD_MAIN_SNS_WLED_UVI80|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>VDD_MAIN_SNS_WLED_UVI80|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = ""                                                          '<Comment>VDD_MAIN_SNS_WLED_UVI80|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>VDD_MAIN_SNS_WLED_UVI80|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>VDD_MAIN_SNS_WLED_UVI80|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>VDD_MAIN_SNS_WLED_UVI80|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>VDD_MAIN_SNS_WLED_UVI80|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>VDD_MAIN_SNS_WLED_UVI80|TestItem|TM_IIL_IIH =</Comment>

    Case 19                                                                    '<Comment>IDAC_OUT_0</Comment>
        SP_PinName = "IDAC_OUT_0"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_0|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_0|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_0|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1043,K1061"                                               '<Comment>IDAC_OUT_0|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_0|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_0|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_0|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_0|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_0|TestItem|TM_IIL_IIH =</Comment>

    Case 20                                                                    '<Comment>IDAC_OUT_1</Comment>
        SP_PinName = "IDAC_OUT_1"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_1|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_1|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_1|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1043,K1060"                                               '<Comment>IDAC_OUT_1|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_1|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_1|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_1|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_1|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_1|TestItem|TM_IIL_IIH =</Comment>

    Case 21                                                                    '<Comment>IDAC_OUT_2</Comment>
        SP_PinName = "IDAC_OUT_2"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_2|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_2|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_2|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1043,K1059"                                               '<Comment>IDAC_OUT_2|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_2|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_2|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_2|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_2|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_2|TestItem|TM_IIL_IIH =</Comment>

    Case 22                                                                    '<Comment>IDAC_OUT_3</Comment>
        SP_PinName = "IDAC_OUT_3"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_3|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_3|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_3|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1043,K1058"                                               '<Comment>IDAC_OUT_3|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_3|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_3|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_3|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_3|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_3|TestItem|TM_IIL_IIH =</Comment>

    Case 23                                                                    '<Comment>IDAC_OUT_4</Comment>
        SP_PinName = "IDAC_OUT_4"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_4|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_4|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_4|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1043,K1057"                                               '<Comment>IDAC_OUT_4|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_4|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_4|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_4|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_4|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_4|TestItem|TM_IIL_IIH =</Comment>

    Case 24                                                                    '<Comment>IDAC_OUT_5</Comment>
        SP_PinName = "IDAC_OUT_5"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_5|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_5|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_5|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1043,K1056"                                               '<Comment>IDAC_OUT_5|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_5|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_5|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_5|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_5|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_5|TestItem|TM_IIL_IIH =</Comment>

    Case 25                                                                    '<Comment>IDAC_OUT_6</Comment>
        SP_PinName = "IDAC_OUT_6"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_6|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_6|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_6|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1041,K1055"                                               '<Comment>IDAC_OUT_6|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_6|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_6|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_6|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_6|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_6|TestItem|TM_IIL_IIH =</Comment>

    Case 26                                                                    '<Comment>IDAC_OUT_7</Comment>
        SP_PinName = "IDAC_OUT_7"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_7|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_7|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_7|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1041,K1054"                                               '<Comment>IDAC_OUT_7|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_7|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_7|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_7|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_7|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_7|TestItem|TM_IIL_IIH =</Comment>

    Case 27                                                                    '<Comment>IDAC_OUT_8</Comment>
        SP_PinName = "IDAC_OUT_8"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_8|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_8|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_8|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1041,K1053"                                               '<Comment>IDAC_OUT_8|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_8|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_8|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_8|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_8|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_8|TestItem|TM_IIL_IIH =</Comment>

    Case 28                                                                    '<Comment>IDAC_OUT_9</Comment>
        SP_PinName = "IDAC_OUT_9"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_9|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_9|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_9|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1041,K1052"                                               '<Comment>IDAC_OUT_9|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_9|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_9|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_9|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_9|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_9|TestItem|TM_IIL_IIH =</Comment>

    Case 29                                                                    '<Comment>IDAC_OUT_10</Comment>
        SP_PinName = "IDAC_OUT_10"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_10|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_10|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_10|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1041,K1051"                                               '<Comment>IDAC_OUT_10|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_10|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_10|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_10|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_10|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_10|TestItem|TM_IIL_IIH =</Comment>

    Case 30                                                                    '<Comment>IDAC_OUT_11</Comment>
        SP_PinName = "IDAC_OUT_11"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_11|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_11|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_11|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1041,K1050"                                               '<Comment>IDAC_OUT_11|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_11|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_11|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_11|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_11|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_11|TestItem|TM_IIL_IIH =</Comment>

    Case 31                                                                    '<Comment>IDAC_OUT_12</Comment>
        SP_PinName = "IDAC_OUT_12"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_12|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_12|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_12|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1039,K1049"                                               '<Comment>IDAC_OUT_12|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_12|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_12|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_12|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_12|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_12|TestItem|TM_IIL_IIH =</Comment>

    Case 32                                                                    '<Comment>IDAC_OUT_13</Comment>
        SP_PinName = "IDAC_OUT_13"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_13|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_13|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_13|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1039,K1048"                                               '<Comment>IDAC_OUT_13|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_13|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_13|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_13|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_13|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_13|TestItem|TM_IIL_IIH =</Comment>

    Case 33                                                                    '<Comment>IDAC_OUT_14</Comment>
        SP_PinName = "IDAC_OUT_14"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_14|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_14|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_14|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1039,K1047"                                               '<Comment>IDAC_OUT_14|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_14|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_14|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_14|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_14|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_14|TestItem|TM_IIL_IIH =</Comment>

    Case 34                                                                    '<Comment>IDAC_OUT_15</Comment>
        SP_PinName = "IDAC_OUT_15"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_15|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_15|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_15|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1039,K1046"                                               '<Comment>IDAC_OUT_15|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_15|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_15|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_15|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_15|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_15|TestItem|TM_IIL_IIH =</Comment>

    Case 35                                                                    '<Comment>IDAC_OUT_16</Comment>
        SP_PinName = "IDAC_OUT_16"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_16|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_16|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_16|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1039,K1045"                                               '<Comment>IDAC_OUT_16|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_16|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_16|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_16|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_16|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_16|TestItem|TM_IIL_IIH =</Comment>

    Case 36                                                                    '<Comment>IDAC_OUT_17</Comment>
        SP_PinName = "IDAC_OUT_17"
        SPLimit = ""                                                           '<Comment>IDAC_OUT_17|SpecificLimit|SPLimit = </Comment> 
        SPIRange = ""                                                          '<Comment>IDAC_OUT_17|SpecificIRange|SPIRange =</Comment> 
        Wait_Time = "0.045"                                                    '<Comment>IDAC_OUT_17|SpecificWaitTime|Wait_Time =</Comment> 
        On_Relay = "K1039,K1044"                                               '<Comment>IDAC_OUT_17|On_Relay|On_Relay =</Comment> 
        Off_Relay = ""                                                         '<Comment>IDAC_OUT_17|Off_Relay|Off_Relay =</Comment> 
        MustDiscnctPins = ""                                                   '<Comment>IDAC_OUT_17|MustDiscnctPins|MustDiscnctPins =</Comment> 
        SPCondPin = ""                                                         '<Comment>IDAC_OUT_17|SpecCondiPin|SPCondPin =</Comment>
        SPCondPinV_I = ""                                                      '<Comment>IDAC_OUT_17|SpecCondiPinVolt_Current|SPCondPinV_I =</Comment> 
        TestItem =  "Both"                                                     '<Comment>IDAC_OUT_17|TestItem|TM_IIL_IIH =</Comment>
    Case Else
             'SP_Leak_Pins_Cond = "Wrong_Enum_Input"
    End Select

    SP_Leak_Pins_Cond= SP_PinName + "&" + SPLimit + "&" + SPIRange + "&" + Wait_Time + "&" + On_Relay + "&" + Off_Relay + "&" + MustDiscnctPins + "&" + SPCondPin + "&" + SPCondPinV_I + "&" + TestItem

End Function


Public Function GenLeakPinDic()
Dim idx As Double
Dim Pin_Num As Double
Dim PinName() As String
Dim Concat As String
Pin_Num = 37
ReDim SPPins(36) As String

Dim Dic_PinName() As String
ReDim Preserve Dic_PinName(Pin_Num)

    For idx = 0 To Pin_Num

    ReDim Preserve PinName(Pin_Num)
    PinName(idx) = Split_Concat(SP_Leak_Pins_Cond(idx), 0) '.SP_Leak_Pins
    
    Dic_PinName(idx) = PinName(idx)
    
    If LeakPinDic.Exists(Dic_PinName(idx)) Then
    Else
       LeakPinDic.Add Dic_PinName(idx), idx     'add key and item to dictionary
    End If
    Next idx

End Function


Public Function SearchDicIdx(PinName As String) As Double
Dim idx As Double
For idx = 0 To 37
    If LeakPinDic.Keys(idx) = PinName Then
    SearchDicIdx = idx
    Exit For
    End If
Next idx
End Function

