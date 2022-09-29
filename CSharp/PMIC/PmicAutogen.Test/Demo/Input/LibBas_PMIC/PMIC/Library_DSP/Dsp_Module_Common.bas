Attribute VB_Name = "Dsp_Module_Common"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.60.70/ with Build Version - 2.23.60.70
'Test Plan:D:\Jeffli\Cota\COTA_A0_TestPlan_0401G.xlsx, MD5=e167926ad80a228efa1541adbfb24269
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:D:\ADC\ProjectAutomation\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :D:\Jeffli\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!


'20181217 evans : for ldo soft start
Public Function CalcSofStartPlusVout(ByVal ResultWave As DSPWave, _
                            ByVal d10or90PrecOfTarget As Double, _
                            ByVal d90or10PrecOfTarget As Double, _
                            ByVal dSampleRate As Double, _
                            ByRef dResultTime As Double, _
                            ByRef dVout As Double _
                            ) As Long
    
    Dim V_10Percentage_wave As New DSPWave
    Dim V_90Percentage_wave As New DSPWave
    Dim start_V As Long
    Dim stop_V As Long
    
    On Error Resume Next
    
    dVout = ResultWave.CalcMaximumValue
    d10or90PrecOfTarget = dVout * 0.1
    d90or10PrecOfTarget = dVout * 0.9
    V_10Percentage_wave = ResultWave.FindIndices(GreaterThanOrEqualTo, d10or90PrecOfTarget)
    V_90Percentage_wave = ResultWave.FindIndices(GreaterThanOrEqualTo, d90or10PrecOfTarget)
    start_V = V_10Percentage_wave.CalcMinimumValue
    stop_V = V_90Percentage_wave.CalcMinimumValue
    dResultTime = Abs((start_V - stop_V) / (dSampleRate))
    
End Function

'20200102 evans : for acore vbat-ibat TTR
Public Function CalcGainOffset_Lo(ByVal sdAvgV As Double, ByVal sdAvgI As Double, sdGain As Double, sdOffset As Double, _
                        Gain_MSB As Double, Gain_LSB As Double, Offset_MSB As Double, Offset_LSB As Double) As Long
    
    On Error Resume Next
    Dim gA3_M As Double
    Dim gA3_M2 As Double
    Dim CA3_M As Double
    Dim CA3_S As Double
    Dim CA3_N As Double
    Dim gA3 As Double
    Dim CA3 As Double
    Dim upperField As Long
    Dim lowerField As Long
    Dim upperFieldOffset As Long
    Dim Offset_Sign As Long
    
    gA3_M = 0.06
    gA3_M2 = sdAvgV - sdAvgI
    If gA3_M2 = 0 Then
        gA3 = -1
        CA3 = -1
    Else
        gA3 = gA3_M / gA3_M2
        CA3_M = 0.03 '30*10^-3
        CA3_S = gA3 * (sdAvgV)
        CA3_N = 0.0375
        CA3 = CA3_M - CA3_S + CA3_N 'CA3_S
    End If
    
    'calcute gain/offset MSB and LSB
    upperField = &HFF00
    lowerField = &HFF
    upperFieldOffset = &H7F00
    Offset_Sign = 0
    If CA3 < 0 Then Offset_Sign = &H80
    
    sdGain = gA3
    sdOffset = CA3
    
    gA3 = gA3 * 1000000000
    CA3 = CA3 * 1000000
    
    Gain_MSB = (gA3 And upperField) / lowerField
    Gain_LSB = gA3 And lowerField
    
    CA3 = Abs(CA3)
    Offset_MSB = (CA3 And upperFieldOffset) / lowerField
    Offset_MSB = Offset_Sign + Offset_MSB
    Offset_LSB = CA3 And lowerField
    
End Function

'20200102 evans : for acore vbat-ibat TTR
Public Function CalcGainOffset_Hi(ByVal sdAvgV As Double, ByVal sdAvgI As Double, sdGain As Double, sdOffset As Double, _
                        Gain_MSB As Double, Gain_LSB As Double, Offset_MSB As Double, Offset_LSB As Double) As Long
    
    On Error Resume Next
    Dim gA3_M As Double
    Dim gA3_M2 As Double
    Dim CA3_M As Double
    Dim CA3_S As Double
    Dim CA3_N As Double
    Dim gA3 As Double
    Dim CA3 As Double
    Dim upperField As Long
    Dim lowerField As Long
    Dim upperFieldOffset As Long
    Dim Offset_Sign As Long
    
    gA3_M = 100 * 0.001
    gA3_M2 = sdAvgV - sdAvgI
    
    gA3 = gA3_M / gA3_M2
    CA3_M = 50 * 0.001
    CA3_S = gA3 * (sdAvgV)
    CA3_N = 62.5 * 0.001
    CA3 = CA3_M - CA3_S + CA3_N
    
    'calcute gain/offset MSB and LSB
    upperField = &HFF00
    lowerField = &HFF
    upperFieldOffset = &H7F00
    Offset_Sign = 0
    If CA3 < 0 Then Offset_Sign = &H80
    
    sdGain = gA3
    sdOffset = CA3
    
    gA3 = gA3 * 1000000000
    CA3 = CA3 * 1000000
    
    Gain_MSB = (gA3 And upperField) / lowerField
    Gain_LSB = gA3 And lowerField
    
    CA3 = Abs(CA3)
    Offset_MSB = (CA3 And upperFieldOffset) / lowerField
    Offset_MSB = Offset_Sign + Offset_MSB
    Offset_LSB = CA3 And lowerField
    
End Function

'20200109 evans : For Acore ADC Gain/Offset
Public Function Best_fit_line_DSP(ByVal x As DSPWave, ByVal AdcCodeOffset As Long, ByVal y As DSPWave, FIT_M As Double, FIT_B As Double, _
                                    Gain_MSB As Double, Gain_LSB As Double, Offset_MSB As Double, Offset_LSB As Double) As Long

    On Error Resume Next
    
    Dim sum_x As Double
    Dim sum_x2 As Double
    Dim sum_xy As Double
    Dim sum_y As Double
    Dim sum_y2 As Double
    
    Dim Index As Long
    Dim temp_x As Double
    Dim temp_x2 As Double
    Dim temp_y As Double
    Dim temp_y2 As Double
    
    Dim FIT_M_U As Double
    Dim FIT_M_D As Double
    Dim FIT_B_U As Double
    Dim FIT_B_D As Double
    
    Dim tmpFIT_M As Double
    Dim tmpFIT_B As Double
    
    Dim dspAdcCode As New DSPWave
    Dim upperField As Long
    Dim lowerField As Long
    Dim upperFieldOffset As Long
    Dim Offset_Sign As Long
    
    Dim arrTemp_x() As Long
    Dim arrTemp_y() As Double
    
    dspAdcCode = x.Select(AdcCodeOffset * y.SampleSize, 1, y.SampleSize)
    sum_y = 0
    sum_x = 0
    temp_x = 0
    temp_y = 0
    sum_x2 = 0
    sum_y2 = 0
    FIT_M_U = 0
    FIT_M_D = 0
    FIT_B_U = 0
    FIT_B_D = 0

'20191119 evans : optimization
    arrTemp_x = dspAdcCode.Data
    arrTemp_y = y.Data
    
    For Index = 0 To y.SampleSize - 1
    
        temp_x = arrTemp_x(Index)
        temp_y = arrTemp_y(Index)
        
        sum_x = sum_x + temp_x
        sum_x2 = sum_x2 + (temp_x * temp_x)
        
        sum_y = sum_y + temp_y
        sum_xy = sum_xy + (temp_x * (temp_y))
        
    Next Index
    
    FIT_M_U = sum_xy * (y.SampleSize) - (sum_x * sum_y) '2*50 - 3*10 = 70 for offline test
    FIT_M_D = sum_x2 * (y.SampleSize) - (sum_x * sum_x)  '4*50 - 3*3 = 191 for offline test
    FIT_M = FIT_M_U / (FIT_M_D) ' 70 / 191 = 0.366 for offline test

    FIT_B_U = sum_y * (sum_x2) - (sum_xy * sum_x) '10*4 - 2*3 = 34 for offline test
    FIT_B_D = sum_x2 * (y.SampleSize) - (sum_x * sum_x) '50*4 - 3*3 = 191 for offline test
    FIT_B = FIT_B_U / (FIT_B_D) ' 34 / 191 = 0.17 for offline test
    
'calcute gain/offset MSB and LSB
    upperField = &HFF00
    lowerField = &HFF
    upperFieldOffset = &H7F00
    Offset_Sign = 0

    tmpFIT_M = FIT_M
    tmpFIT_M = tmpFIT_M * 1000000
    Gain_MSB = (tmpFIT_M And upperField) / lowerField
    Gain_LSB = tmpFIT_M And lowerField
    
    tmpFIT_B = FIT_B
    If tmpFIT_B < 0 Then
        Offset_Sign = &H80
    End If
    
    tmpFIT_B = Abs(tmpFIT_B) * 1000000
    Offset_MSB = (tmpFIT_B And upperFieldOffset) / lowerField
    Offset_MSB = Offset_Sign + Offset_MSB
    Offset_LSB = tmpFIT_B And lowerField
    
End Function
