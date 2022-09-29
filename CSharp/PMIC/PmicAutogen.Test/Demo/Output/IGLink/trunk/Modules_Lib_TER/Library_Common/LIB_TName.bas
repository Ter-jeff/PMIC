Attribute VB_Name = "LIB_TName"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!

Option Explicit
Public gVDD       As Double

''''Using below, it will need to import "LIB_TestName.bas/LibTestName.cls/LibTestNameGrp7.cls"
''Public UniFormTestName As New LibTestName
'====================================================================
'   -- Version --
'   Date      |  Auther  |  version
'   20180830  |  Neo     |  1.00
'   20180906  |  Neo     |  1.01
'  _____________________________________________________________
'   Note:
'  --------------------------------------------
'   Grp1 ~ 4 with the same rull:
'       if Grp1 = ""   -> Auto fill in Instance Name 1st column
'          Grp1 <> ""  -> Fill the customize name
'
'   if Grp8 = "X"      -> refer to Grp7,
'                         check Grp7 String is TrimValue or Code
'
'   if Grp9 = "X"      -> refer to globale variable of VDD
'
'====================================================================
'
'
''*******************************************************************
''******************        Reference          **********************
''*******************************************************************
'' Test Naming Convention
'-------------------------------------------------
'  Grp1_Name = WhichBuck & "_"   'TestBlock
'  Grp2_Name = WhichPhase & "_"  'PhaseNum
'  Grp3_Name = "LS_"             'TestMode
'  Grp4_Name = "CSA_"            'SubTestMode
'  Grp5_Name = "X_"              'SubTestCondition
'  Grp6_Name = "X_"              'TBD
'  Grp7_Name = "PostBurnCode_"   'TrimConditions
'  Grp8_Name = "X_"              'LinkNum
'  Grp9_Name = VDDSupply         'VDDLevel
'  Grp10_Name = "ToggleDTB_"     'MeasureType
'  Grp11_Name = "X_"             'TBD
'  Grp12_Name = "X"              'TBD
'____________________________________________________________________
'____________________________________________________________________
'
'**************************  Enum  *****************************


Public g_Number_real_bak As Double
Public gVDD_BUCK  As Double
Public Enum Group7Enum
    TName_PreTrim = 1
    TName_PreTrimCode
    TName_PreTrimDelta

    TName_PostTrim
    TName_PostTrimCode
    TName_PostTrimDelta

    TName_TweakTrim
    TName_TweakTrimCode
    TName_TweakTrimDelta

    TName_Swp5
    TName_SweepCodes

    TName_FinalTrim
    TName_FinalTrimCode
    TName_FinalTrimDelta
    TName_FinalTrimGoNogo

    TName_TrimLink
    TName_TrimTarget

    TName_PostBurn
    TName_PostBurnCode
    TName_PostBurnTarget

    TName_PostBurnGNG
    TName_PostBurnGNGCHAR

    '    TName_MeasR                    ''''20190923 remove for group10 tname------Henry
    TName_NonTrimItem

    TName_NoEntry

    TName_OTP_Addr
    TName_OTP_X

    TName_CodeSweep
    TName_PostBurnSweepCode
    TName_PostBurnSweepShift

    '--- you can add more condition at below       ---
    '--- MUST update function <Group7EnumToString> ---

End Enum

Public Enum Group10Enum
    TName_MeasI = 1
    TName_MeasV
    TName_MeasF
    TName_MeasG
    TName_CalcR
    TName_CalcC_FW

    TName_MeasR

    TName_MeasV_Delta
    TName_MeasV_Norm
    TName_MeasV_FW

    TName_MeasI_Delta
    TName_MeasI_Norm
    TName_MeasT_Delta
    TName_MeasT_Norm

    TName_MeasT
    TName_CalcT
    TName_ToggleDTB
    TName_MeasPW
    TName_MeasTime
    TName_MeasT_Delta_FW
    TName_MeasT_Norm_FW
    TName_MeasT_FW

    TName_MeasCode
    TName_MeasLSB

    TName_None
    TName_OTP_Grp10_Actual
    TName_OTP_Grp10_Expected
    TName_OTP_Grp10_Match

    TName_MeasTemp


    '--- you can add more condition at below       ---
    '--- MUST update function <Group10EnumToString> ---

End Enum
'**************************  Enum End  *****************************
'**************************  Enum To String *****************************
Public Function Group7EnumToString(EnumInput As Group7Enum) As String
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Group7EnumToString"

    Select Case EnumInput
        Case TName_PreTrim: Group7EnumToString = "PreTrim"
        Case TName_PreTrimCode: Group7EnumToString = "PreTrimCode"
        Case TName_PreTrimDelta: Group7EnumToString = "PreTrimDelta"

        Case TName_PostTrim: Group7EnumToString = "PostTrim"
        Case TName_PostTrimCode: Group7EnumToString = "PostTrimCode"
        Case TName_PostTrimDelta: Group7EnumToString = "PostTrimDelta"

        Case TName_TweakTrim: Group7EnumToString = "TweakTrim"
        Case TName_TweakTrimCode: Group7EnumToString = "TweakTrimCode"
        Case TName_TweakTrimDelta: Group7EnumToString = "TweakTrimDelta"

        Case TName_FinalTrim: Group7EnumToString = "FinalTrim"
        Case TName_FinalTrimCode: Group7EnumToString = "FinalTrimCode"
        Case TName_FinalTrimDelta: Group7EnumToString = "FinalTrimDelta"

        Case TName_TrimLink: Group7EnumToString = "TrimLink"
        Case TName_TrimTarget: Group7EnumToString = "TrimTarget"
        Case TName_Swp5: Group7EnumToString = "Swp5"

        Case TName_FinalTrimGoNogo: Group7EnumToString = "FinalTrimGNG"
        Case TName_SweepCodes: Group7EnumToString = "SweepCodes"

        Case TName_PostBurn: Group7EnumToString = "PostBurn"
        Case TName_PostBurnCode: Group7EnumToString = "PostBurnCode"
        Case TName_PostBurnTarget: Group7EnumToString = "PostBurnTarget"

        Case TName_PostBurnGNG: Group7EnumToString = "PostBurnGNG"
        Case TName_PostBurnGNGCHAR: Group7EnumToString = "PostBurnGNGCHAR"

        Case TName_NonTrimItem: Group7EnumToString = "P"
        Case TName_OTP_Addr: Group7EnumToString = "Addr"
        Case TName_OTP_X: Group7EnumToString = "X"

        Case TName_CodeSweep: Group7EnumToString = "CodeSweep"
        Case TName_PostBurnSweepCode: Group7EnumToString = "PostBurnSweepCode"
        Case TName_PostBurnSweepShift: Group7EnumToString = "PostBurnSweepShift"

        Case Else
            Group7EnumToString = "Wrong_Enum_Input"
    End Select

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Group10EnumToString(EnumInput As Group10Enum) As String
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Group10EnumToString"

    Select Case EnumInput

        Case TName_MeasI: Group10EnumToString = "MeasI"
        Case TName_MeasV: Group10EnumToString = "MeasV"
        Case TName_MeasF: Group10EnumToString = "MeasF"
        Case TName_MeasG: Group10EnumToString = "MeasG"
        Case TName_CalcR: Group10EnumToString = "CalcR"
        Case TName_CalcC_FW: Group10EnumToString = "CalcC-FW"

        Case TName_MeasR: Group10EnumToString = "MeasR"
        Case TName_MeasTime: Group10EnumToString = "MeasTime"


        Case TName_MeasV_Norm: Group10EnumToString = "MeasV-Norm"
        Case TName_MeasV_Delta: Group10EnumToString = "MeasV-Delta"
        Case TName_MeasV_FW: Group10EnumToString = "MeasV-FW"

        Case TName_MeasI_Norm: Group10EnumToString = "MeasI-Norm"
        Case TName_MeasI_Delta: Group10EnumToString = "MeasI-Delta"
        Case TName_MeasT_Norm: Group10EnumToString = "MeasT-Norm"
        Case TName_MeasT_Delta: Group10EnumToString = "MeasT-Delta"
        Case TName_MeasT_Norm_FW: Group10EnumToString = "MeasT-Norm-FW"
        Case TName_MeasT_Delta_FW: Group10EnumToString = "MeasT-Delta-FW"
        Case TName_MeasT_FW: Group10EnumToString = "MeasT-FW"

        Case TName_MeasTemp: Group10EnumToString = "MeasTemp"

        Case TName_MeasT: Group10EnumToString = "MeasT"
        Case TName_CalcT: Group10EnumToString = "CalcT"
        Case TName_ToggleDTB: Group10EnumToString = "ToggleDTB"

        Case TName_MeasCode: Group10EnumToString = "MeasCode"
        Case TName_MeasLSB: Group10EnumToString = "MeasLSB"

        Case TName_None: Group10EnumToString = "X"
        Case TName_OTP_Grp10_Actual: Group10EnumToString = "Actual"
        Case TName_OTP_Grp10_Expected: Group10EnumToString = "Expected"
        Case TName_OTP_Grp10_Match: Group10EnumToString = "Match"
        Case Else
            Group10EnumToString = "Wrong_Enum_Input"

    End Select

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'**************************  Enum To String End *****************************

Public Function Group5_judge(input_grp5 As Variant, input_grp10 As String) As String
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Group5_judge"

    ''''-----/**************************************************************************
    ''''-----/**************************************************************************
    ''''-----/**************************************************************************

    '    Dim Grp10 As String  ' this value just use in number judge , when string this value doesn't need
    'Dim input_grp5 As Variant   'user input
    Dim Grp5_double As Variant    'need add in function
    Dim Grp5_srtArr() As String    'need add in function
    Dim output_Grp5 As String    ' we output

    '    input_grp5 = 0.0002
    '    Grp5 = "0.0002"
    '    Grp5 = "0.0002mA"
    '    Grp5 = "1.5Mhz"

    '     Grp10 = "TN_MeasT"
    '     Grp10 = "TN_MeasI"
    '     Grp10 = "TN_MeasV"
    '     Grp10 = "TN_MeasF"
    '     Grp10 = "TN_MeasR"
    '     output_Grp5 = "XXXX"


    '------------------------------------- input is Numeric
    If IsNumeric(input_grp5) = True Then
        If InStr(input_grp10, "MeasT") > 0 Then       'MeasT, unit: sec
            If input_grp5 < 1 * ns Then
                input_grp5 = input_grp5 / (1 * ps)
                output_Grp5 = CStr(input_grp5) & "ps"
            ElseIf input_grp5 < 1 * us Then
                input_grp5 = input_grp5 / (1 * ns)
                output_Grp5 = CStr(input_grp5) & "ns"
            Else
                input_grp5 = input_grp5 / (1 * us)
                output_Grp5 = CStr(input_grp5) & "us"
            End If

        ElseIf InStr(input_grp10, "CalcT") > 0 Then    'CalcT, unit: uA/us
            output_Grp5 = CStr(input_grp5) & "uA/us"

        ElseIf InStr(input_grp10, "R") > 0 Then        'R, unit: ohm
            output_Grp5 = CStr(input_grp5) & "ohm"

            If Abs(input_grp5) > 1 * k Then
                input_grp5 = input_grp5 / 1000
                output_Grp5 = CStr(input_grp5) & "kohm"
            End If

            If Abs(input_grp5) < 1 Then
                input_grp5 = input_grp5 * 1000
                output_Grp5 = CStr(input_grp5) & "mohm"
            End If

        ElseIf InStr(input_grp10, "I") > 0 Then         'I, unit: A
            output_Grp5 = CStr(input_grp5) & "A"

            If Abs(input_grp5) < 1 And Abs(input_grp5) > 0.000999 Then
                input_grp5 = input_grp5 * 1000
                output_Grp5 = CStr(input_grp5) & "mA"
            End If

            If Abs(input_grp5) < 0.001 Then
                input_grp5 = input_grp5 * 1000000
                output_Grp5 = CStr(input_grp5) & "uA"
            End If

        ElseIf InStr(input_grp10, "V") > 0 Or InStr(input_grp10, "DTB") > 0 Then     'V,  unit: V
            output_Grp5 = CStr(input_grp5) & "V"

            If Abs(input_grp5) < 1 Then
                input_grp5 = input_grp5 * 1000
                output_Grp5 = CStr(input_grp5) & "mV"
            End If

        ElseIf InStr(input_grp10, "F") > 0 Then         'F, unit: Hz
            output_Grp5 = CStr(input_grp5) & "Hz"

            If Abs(input_grp5) > 1 * k Then
                input_grp5 = input_grp5 / 1 * k
                output_Grp5 = CStr(input_grp5) & "kHz"
            End If

            If Abs(input_grp5) > 1 * m Then
                input_grp5 = input_grp5 / 1 * m
                output_Grp5 = CStr(input_grp5) & "MHz"
            End If

        ElseIf InStr(input_grp10, "G") > 0 Then          'G, unit:
            output_Grp5 = CStr(input_grp5) & ""

        Else
            'if Grp10 undefine
            'input_grp5 = "Undefine_" & CStr(input_grp5)
            input_grp5 = CStr(input_grp5) & ""

        End If    ' end of Grp5 is number judge

    Else
        '------------------------------------------- input is String

        If InStr(input_grp5, "A") > 0 Then  '------ current
            If InStr(input_grp5, "m") > 0 Then
                Grp5_srtArr = Split(input_grp5, "mA")
                Grp5_double = CDbl(Grp5_srtArr(0))

                If Abs(Grp5_double) < 1 And Abs(Grp5_double) > 0.000999 Then
                    Grp5_double = Grp5_double * 1000
                    output_Grp5 = CStr(Grp5_double) & "uA"
                ElseIf Abs(Grp5_double) < 0.001 Then
                    Grp5_double = Grp5_double * 1000000
                    output_Grp5 = CStr(Grp5_double) & "nA"
                Else
                    output_Grp5 = CStr(Grp5_double) & "mA"
                End If

            ElseIf InStr(input_grp5, "u") > 0 Then
                Grp5_srtArr = Split(input_grp5, "uA")
                Grp5_double = CDbl(Grp5_srtArr(0))

                If Abs(Grp5_double) < 1 And Abs(Grp5_double) > 0.000999 Then
                    Grp5_double = Grp5_double * 1000
                    output_Grp5 = CStr(Grp5_double) & "nA"
                Else
                    output_Grp5 = CStr(Grp5_double) & "uA"
                End If

            ElseIf InStr(input_grp5, "n") > 0 Then
                Grp5_srtArr = Split(input_grp5, "nA")
                Grp5_double = CDbl(Grp5_srtArr(0))

                If Abs(Grp5_double) < 1 And Abs(Grp5_double) > 0.000999 Then
                    Grp5_double = Grp5_double * 1000
                    output_Grp5 = CStr(Grp5_double) & "pA"
                Else
                    output_Grp5 = CStr(Grp5_double) & "nA"
                End If

            Else
                Grp5_srtArr = Split(input_grp5, "A")
                Grp5_double = CDbl(Grp5_srtArr(0))

                If Abs(Grp5_double) < 1 And Abs(Grp5_double) > 0.000999 Then
                    Grp5_double = Grp5_double * 1000
                    output_Grp5 = CStr(Grp5_double) & "mA"
                ElseIf Abs(Grp5_double) < 0.001 Then
                    Grp5_double = Grp5_double * 1000000
                    output_Grp5 = CStr(Grp5_double) & "uA"
                Else
                    output_Grp5 = CStr(Grp5_double) & "A"
                End If

            End If  ' end of Current string judge

        ElseIf InStr(input_grp5, "V") > 0 Then  '------ voltage
            If InStr(input_grp5, "m") > 0 Then
                Grp5_srtArr = Split(input_grp5, "mV")
                Grp5_double = CDbl(Grp5_srtArr(0))

                If Abs(Grp5_double) < 1 And Abs(Grp5_double) > 0.000999 Then
                    Grp5_double = Grp5_double * 1000
                    output_Grp5 = CStr(Grp5_double) & "uV"
                ElseIf Abs(Grp5_double) < 0.001 Then
                    Grp5_double = Grp5_double * 1000000
                    output_Grp5 = CStr(Grp5_double) & "nV"
                Else
                    output_Grp5 = CStr(Grp5_double) & "mV"
                End If

            ElseIf InStr(input_grp5, "u") > 0 Then
                Grp5_srtArr = Split(input_grp5, "uV")
                Grp5_double = CDbl(Grp5_srtArr(0))

                If Abs(Grp5_double) < 1 And Abs(Grp5_double) > 0.000999 Then
                    Grp5_double = Grp5_double * 1000
                    output_Grp5 = CStr(Grp5_double) & "nV"
                Else
                    output_Grp5 = CStr(Grp5_double) & "uV"
                End If

            ElseIf InStr(input_grp5, "n") > 0 Then
                Grp5_srtArr = Split(input_grp5, "nV")
                Grp5_double = CDbl(Grp5_srtArr(0))

                If Abs(Grp5_double) < 1 And Abs(Grp5_double) > 0.000999 Then
                    Grp5_double = Grp5_double * 1000
                    output_Grp5 = CStr(Grp5_double) & "pV"
                Else
                    output_Grp5 = CStr(Grp5_double) & "nV"
                End If

            Else
                Grp5_srtArr = Split(input_grp5, "V")
                Grp5_double = CDbl(Grp5_srtArr(0))
                If Abs(Grp5_double) < 1 And Abs(Grp5_double) > 0.000999 Then
                    Grp5_double = Grp5_double * 1000
                    output_Grp5 = CStr(Grp5_double) & "mV"
                ElseIf Abs(Grp5_double) < 0.001 Then
                    Grp5_double = Grp5_double * 1000000
                    output_Grp5 = CStr(Grp5_double) & "uV"
                Else
                    output_Grp5 = CStr(Grp5_double) & "V"
                End If
            End If   ' end of Voltage string judge

        Else
            ' when group 5 use unknow string unit
            output_Grp5 = CStr(input_grp5)
        End If    ' end of Grp5 string judge
    End If    ' end of all Grp5 define

    Group5_judge = output_Grp5

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'****************************************************************************
'                              Main
'****************************************************************************
Public Function TNameCombine(Optional Grp1 As String = "X", _
                             Optional Grp2 As String = "X", _
                             Optional Grp3 As String = "X", _
                             Optional Grp4 As String = "X", _
                             Optional Grp5 As Variant = "X", _
                             Optional Grp6 As String = "X", _
                             Optional Grp7 As Group7Enum, _
                             Optional Grp8 As String = "X", _
                             Optional Grp9 As Variant = "X", _
                             Optional Grp10 As Group10Enum, _
                             Optional Grp11 As Variant = "X", _
                             Optional Grp12 As Variant = "X", _
                             Optional TestNumber As Double = 99) As String
    '-----------------------------------------------------------

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "TNameCombine"

    Dim GrpArr(1 To 12) As String
    Dim VDD_MAIN_PIN As String: VDD_MAIN_PIN = "VDD_MAIN_UVI80"
    Dim InsNameArr() As String
    InsNameArr = Split(TheExec.DataManager.InstanceName, "_")

    '------------------------------ Group 1, 2, 3, 4: TestBlock
    If Grp1 = "X" Then
        If UBound(InsNameArr) > -1 Then
            If InsNameArr(0) <> "" Then GrpArr(1) = UCase(InsNameArr(0))
        Else
            GrpArr(1) = "X"
        End If
    Else
        If Grp1 = "" Then
            GrpArr(1) = "X"
        Else
            GrpArr(1) = UCase(Grp1)
        End If
    End If

    If Grp2 = "X" Then
        If UBound(InsNameArr) > 0 Then
            If InsNameArr(1) <> "" Then GrpArr(2) = InsNameArr(1)
        Else
            GrpArr(2) = "X"
        End If
    Else
        If Grp2 = "" Then
            GrpArr(2) = "X"
        Else
            GrpArr(2) = Grp2
        End If
    End If

    If Grp3 = "X" Then
        If UBound(InsNameArr) > 1 Then
            If InsNameArr(2) <> "" Then GrpArr(3) = InsNameArr(2)
        Else
            GrpArr(3) = "X"
        End If
    Else
        If Grp3 = "" Then
            GrpArr(3) = "X"
        Else
            GrpArr(3) = Grp3
        End If
    End If

    If Grp4 = "X" Then
        If UBound(InsNameArr) > 2 Then
            If InsNameArr(3) <> "" Then GrpArr(4) = InsNameArr(3)
        Else
            GrpArr(4) = "X"
        End If
    Else
        If Grp4 = "" Then
            GrpArr(4) = "X"
        Else
            GrpArr(4) = Grp4
        End If
    End If

    '------------------------------ Group5: SubTestCondition
    '------------------------------         judge after Group10

    '------------------------------ Group6
    GrpArr(6) = Grp6

    '------------------------------ Group7: TrimConditions
    If Grp7 = 0 Then
        GrpArr(7) = "X"  'no input, set to "X"
    Else
        GrpArr(7) = Group7EnumToString(Grp7)
    End If

    '------------------------------ Group8: LinkNum
    If Grp8 = "X" Then
        If InStr(UCase(GrpArr(7)), "SWEEPCODES") > 0 Then
            GrpArr(8) = "xxx"
        ElseIf InStr(UCase(GrpArr(7)), "GNG") > 0 Then
            GrpArr(8) = "X"
        ElseIf InStr(UCase(GrpArr(7)), "TRIM") > 0 Or _
               InStr(UCase(GrpArr(7)), "BURN") > 0 Then
            GrpArr(8) = "T"

        Else
            GrpArr(8) = "X"
        End If

    Else
        GrpArr(8) = Grp8
    End If

    '------------------------------ Group9: VDDLevel

    If Grp9 = "X" Then
        '        Dim gVDD As Double  '<--------- should replace to global VAR
        gVDD = TheHdw.DCVI.Pins(VDD_MAIN_PIN).Voltage
        GrpArr(9) = "VDD" & Format(gVDD, "0.0") & "V"
    Else
        If IsNumeric(Grp9) Then
            GrpArr(9) = "VDD" & Round(Grp9, 2) & "V"
        Else
            If Grp9 = "" Then
                GrpArr(9) = "X"
            Else
                GrpArr(9) = Grp9
            End If
        End If
    End If

    '------------------------------ Group10: MeasureType
    If Grp10 = 0 Then
        GrpArr(10) = "X"  'no input, set to "X"
    Else
        GrpArr(10) = Group10EnumToString(Grp10)
    End If

    '------------------------------ Group5: SubTestCondition
    If Grp5 = "X" Then
        GrpArr(5) = Grp5
    Else
        If IsNumeric(Grp5) Then
            GrpArr(5) = Group5_judge(Grp5, GrpArr(10))
        Else
            GrpArr(5) = Grp5
        End If
    End If

    '------------------------------ Group11
    If Grp11 = "X" Then
        GrpArr(11) = "X"  'no input, set to "X"
    Else
        GrpArr(11) = Grp11
        '        If IsNumeric(Grp11) Then
        '            GrpArr(11) = "MIN-" & CStr(Round(Grp11, 2)) & "V"
        '        Else
        '            GrpArr(11) = Grp11
        '        End If
    End If

    '------------------------------ Group12
    If Grp12 = "X" Then
        'GrpArr(12) = "X" + "_" + g_sDftType 'no input, set to "X"
        GrpArr(12) = "X"    'no input, set to "X"
    Else
        'GrpArr(12) = Grp12 + "_" + g_sDftType
        GrpArr(12) = Grp12
        '        If IsNumeric(Grp12) Then
        '            GrpArr(12) = "MAX-" & CStr(Round(Grp12, 2)) & "V"
        '        Else
        '            GrpArr(12) = Grp12
        '        End If
    End If


    '------------------------------ TestNumber
    If TestNumber <> 99 Then

        Dim TestNumber_Final As Double



        Select Case Grp7
            Case TName_PreTrim: TestNumber_Final = TestNumber + 1

            Case TName_PreTrimCode: TestNumber_Final = TestNumber + 0
            Case TName_PostTrim: TestNumber_Final = TestNumber + 4

            Case TName_PostTrimCode: TestNumber_Final = TestNumber + 3
            Case TName_FinalTrim: TestNumber_Final = TestNumber + 7

            Case TName_FinalTrimCode: TestNumber_Final = TestNumber + 6
            Case TName_PostBurn: TestNumber_Final = TestNumber + 9

            Case TName_PostBurnCode: TestNumber_Final = TestNumber + 8


            Case TName_PreTrimDelta:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_PostTrimDelta:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_TweakTrim:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_TweakTrimCode:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_TweakTrimDelta:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_FinalTrimDelta:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_TrimLink:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_TrimTarget:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_PostBurnTarget:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_PostBurnGNG:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_PostBurnGNGCHAR:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
            Case TName_NonTrimItem:
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1

            Case Else
                TestNumber_Final = g_Number_real_bak
                g_Number_real_bak = g_Number_real_bak + 1
        End Select

        For Each g_Site In TheExec.Sites
            TheExec.Sites(g_Site).TestNumber = TestNumber_Final
        Next g_Site

    End If    ' end of test number modify


    '--------------------------------------------
    '------------------------------ Final combine
    Dim Final_TNam As String

    Final_TNam = Join(GrpArr, "_")

    Dim strTestNameCheckArray() As String

    '    strTestNameCheckArray = Split(Final_TNam, "_")
    '
    '
    '    End If


    TNameCombine = Final_TNam

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "TNameCombine fail !"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
'****************************************************************************
'                          Main   End
'****************************************************************************
'Public Function TNAMEtest()
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "TNAMEtest"
'
'    Dim final     As String
'
'    final = TNameCombine(, , , , "0.05mA", , , , 3.123456, TName_MeasI)
'    Debug.Print final
'
'    Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

''2019_1213    Duplicate function definition in LIB_TestName module, mask it for bypassing compiler error
''Public Sub RunTimeError(funcName As String)
''On Error GoTo ErrHandler
''
''''''---------------------------This function from LIB_TestName
''    ' Sanity clause
''    If TheExec Is Nothing Then
''        MsgBox "IG-XL in not running!  Error encountered in Exec Interpose Function " + funcName + vbCrLf + _
 ''            "VBT Error # " + Trim$(CStr(err.Number)) + ": " + err.Description
''        Exit Sub
''    End If
''    TheExec.Datalog.WriteComment "Error encountered in Function::" + funcName
''
''Exit Sub
''ErrHandler:
''     If AbortTest Then Exit Sub Else Resume Next
''End Sub



''2019_1213    Duplicate function definition in LIB_TestName module, mask it for bypassing compiler error

'''Public Function GetTNameTemplate(Optional forceVal As String = "X", Optional InstName As String) As String
'''''''---------------------------This function from LIB_TestName
'''    On Error GoTo ErrHandler
'''    Dim funcName As String:: funcName = "GetTNameTemplate"
'''    Dim sKey As String
'''    Dim sTemplateName As String
'''
'''    Exit Function
'''ErrHandler:
'''    RunTimeError funcName
'''    If AbortTest Then Exit Function Else Resume Next
'''End Function
