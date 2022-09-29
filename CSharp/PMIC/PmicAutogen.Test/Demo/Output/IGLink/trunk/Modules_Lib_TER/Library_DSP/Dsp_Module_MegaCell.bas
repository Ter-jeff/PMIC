Attribute VB_Name = "Dsp_Module_MegaCell"
Option Explicit


'20200102 evans : for acore vbat-ibat TTR
Public Function CalcAverageData(ByVal Data0 As DSPWave, ByVal Data1 As DSPWave, ByVal Data2 As DSPWave, ByVal Data3 As DSPWave, ByVal Data4 As DSPWave, _
                                ByVal Data5 As DSPWave, ByVal Data6 As DSPWave, ByVal Data7 As DSPWave, ByVal Data8 As DSPWave, ByVal Data9 As DSPWave, _
                                ByVal Data10 As DSPWave, ByVal Data11 As DSPWave, ByVal Data12 As DSPWave, ByVal Data13 As DSPWave, ByVal Data14 As DSPWave, _
                                ByVal Data15 As DSPWave, sdAvgResult As Double) As Long
    
    On Error Resume Next
    Dim TmpArray() As Long
    
    TmpArray = Data0.Data
    sdAvgResult = TmpArray(2) + TmpArray(6)
    TmpArray = Data1.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data2.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data3.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data4.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data5.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data6.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data7.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data8.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data9.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data10.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data11.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data12.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data13.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data14.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    TmpArray = Data15.Data
    sdAvgResult = sdAvgResult + TmpArray(2) + TmpArray(6)
    
    sdAvgResult = sdAvgResult / 32
    
End Function



  
