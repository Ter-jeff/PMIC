Attribute VB_Name = "DSP_LIB_HardIP"

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

Public Function Measure_DutyCycle(ByVal Wave1 As DSPWave, ByRef dutycycle As Double, ByRef dspStatus As Long) As Long

''Dim Rj As Double
''Dim DDj As Double
''Dim MeasuredUI As Double

Dim RJ As Double
Dim DDJ As Double
Dim MeasuredUI As Double



Call Wave1.MeasureJitter(RJ, DDJ, MeasuredUI, dspStatus)

dutycycle = 100# * ((MeasuredUI - DDJ) / (2 * MeasuredUI))


TheHdw.Digital.Jitter.FileExport Wave1, ".\LPDPADCJitter1"




End Function

Public Function trim_calc(CapWave As DSPWave, outWave As DSPWave) As Long
Dim i As Integer

    'calculate trim code
'        OutWave.Element(0) = 12
    outWave.CreateConstant 0, 5
    
        For i = 0 To 4 'CapWave.CountElements
       
        If (CapWave.Element(i) < 274.42) Then
         outWave.Element(i) = 0
        Else
         outWave.Element(i) = (73.7579 - 10.2598 * Sqr(CapWave.Element(i) - 274.42)) ' update 20130916 ccwu
        End If
        
        'OutWave.Element(i) = (73.134 - 10.153 * Sqr(CapWave.Element(i) - 272.32))  ' 25C
        
        'OutWave.Element(i) = (73.7579 - 10.2598 * Sqr(CapWave.Element(i) - 274.42)) ' update 20130916 ccwu
        'OutWave.Element(i) = 15
        'OutWave.Element(i) = (92.8819 - 6.4957 * Sqr(CapWave.Element(i) - 786.95)) ' 125C
        
            If outWave.Element(i) > 31 Then
                outWave.Element(i) = 31
            ElseIf outWave.Element(i) < 0 Then
                outWave.Element(i) = 0
            End If
        
        Next i
    
End Function


Public Function trim_conveter(InWave As DSPWave, outWave As DSPWave) As Long
Dim i As Long, j As Long
Dim in_temp As Long

    outWave.CreateConstant 0, 3
    
     outWave.Element(0) = InWave.Element(2)
     outWave.Element(1) = InWave.Element(3)
     outWave.Element(2) = InWave.Element(4)
    
''    For j = 0 To 1
''      in_temp = inWave.Element(j)
''        For i = 0 To 4
''          OutWave.Element(j) = in_temp Mod 2
''          in_temp = Int(in_temp / 2)
''        Next i
''    Next j
End Function


Public Function MeasureJitter_US10G(ByVal wave As DSPWave, DDJ As Double, RJ As Double, Meas_UI As Double, dspStatus As Long, ByRef dutycycle As Double, ByRef freq As Double) As Long
'used to be MeasureJitter_vb
    Dim PWHigh As Double
    Dim PWLow As Double
    Dim MeasuredPeriod As Double
   ' Dim dspStatus As Long
   ' Dim dspStatus1 As Long
    Dim UI As Double

    Call wave.SerialMeasureJitter(RJ, DDJ, Meas_UI, dspStatus)
    Call wave.measuretime(PWHigh, PWLow, MeasuredPeriod, dspStatus)
    dutycycle = PWHigh / MeasuredPeriod * 100
    If MeasuredPeriod <> 0 Then
       freq = 1 / MeasuredPeriod
    Else
       freq = 0
    End If
    End Function

Public Function Measure_Eye_US10G(ByVal wave As DSPWave, UI As Double, RJ As Double, DDJ As Double, RiseTime As Double, _
    FallTime As Double, EarlyLow As Double, LateLow As Double, EarlyMid As Double, LateMid As Double, EarlyHigh As Double, _
    LateHigh As Double, dspStatus As Long) As Long
'used to be analyzeEye_vb
    Call wave.SerialMeasureEye(RJ, DDJ, UI, RiseTime, FallTime, EarlyLow, LateLow, EarlyMid, LateMid, EarlyHigh, LateHigh, dspStatus)
End Function

Public Function Measure_DutyCycle_Freq_US10G(ByVal wave As DSPWave, dutycycle As Double, freq As Double) As Long
    Dim RJ        As Double
    Dim DDJ       As Double
    Dim UI        As Double
    Dim dspStatus As Long

    Call wave.SerialMeasureJitter(RJ, DDJ, UI, dspStatus)

    dutycycle = 100# * (UI - DDJ) / (2 * UI)
    freq = 1 / (2 * UI)
    
End Function

Public Function LoopEyeMeas(ByVal wave As DSPWave, RJ As Double, DDJ As Double, Tj As Double, measUI As Double, Tr As Double, _
            Tf As Double, Eye50 As Double, Eye20 As Double, Eye80 As Double, retStatus As Long) As Long
        
    Dim EarlyLow As Double
    Dim LateLow As Double
    Dim EarlyMid As Double
    Dim LateMid As Double
    Dim EarlyHi As Double
    Dim LateHi As Double
    Dim RiseTime As Double
    Dim FallTime As Double
        
    Call wave.MeasureEye(RJ, DDJ, measUI, RiseTime, FallTime, EarlyLow, LateLow, EarlyMid, LateMid, EarlyHi, LateHi, retStatus)

    Call TheHdw.Digital.Jitter.FileExport(wave, "Raw Eye Data")
    
    Eye20 = EarlyLow + measUI - LateLow
    Eye50 = EarlyMid + measUI - LateMid
    Eye80 = EarlyHi + measUI - LateHi
    
End Function

Public Function duty_freq_meas(Captured As DSPWave, ByRef dutycycle As Double, ByRef freq As Double) As Long
    Dim PWHigh As Double
    Dim PWLow As Double
    Dim MeasuredPeriod As Double
    Dim dspStatus As Long
    
    
    Call Captured.measuretime(PWHigh, PWLow, MeasuredPeriod, dspStatus)
    dutycycle = PWHigh / MeasuredPeriod * 100
    If MeasuredPeriod <> 0 Then
       freq = 1 / MeasuredPeriod
    Else
       freq = 0
    End If
End Function


Public Function pulseMeas(ByVal wave As DSPWave, PWHigh As Double, PWLow As Double, Period As Double, Status As Long) As Long
        

        
    Call wave.measuretime(PWHigh, PWLow, Period, Status)
    
End Function
    
Public Function LoopJitterMeas(ByVal wave As DSPWave, RJ As Double, DDJ As Double, measUI As Double, retStatus As Long, ByRef dutycycle As Double, ByRef freq As Double) As Long
        
    Dim PWHigh As Double
    Dim PWLow As Double
    Dim MeasuredPeriod As Double
    Dim dspStatus As Long
        
    Call wave.MeasureJitter(RJ, DDJ, measUI, retStatus)

    Call TheHdw.Digital.Jitter.FileExport(wave, "Raw Jitter Data")
    
    Call wave.measuretime(PWHigh, PWLow, MeasuredPeriod, dspStatus)
    
    dutycycle = PWHigh / MeasuredPeriod * 100
    If MeasuredPeriod <> 0 Then
       freq = 1 / MeasuredPeriod
    Else
       freq = 0
    End If
End Function

Public Function DDR_LoopJitterMeas(ByVal wave As DSPWave, RJ As Double, DDJ As Double, measUI As Double, retStatus As Long, ByRef dutycycle As Double, ByRef freq As Double, _
                                                     ByRef PWHigh As Double, ByRef PWLow As Double, ByRef MeasuredPeriod As Double) As Long
        
''    Dim PWHigh As Double
''    Dim PWLow As Double
''    Dim MeasuredPeriod As Double
    Dim dspStatus As Long
        
''    Call wave.MeasureJitter(RJ, DDJ, measUI, retStatus)
''
''    Call TheHdw.Digital.Jitter.FileExport(wave, "Raw Jitter Data")
    
    Call wave.measuretime(PWHigh, PWLow, MeasuredPeriod, dspStatus)
    
    dutycycle = PWHigh / MeasuredPeriod * 100
    If MeasuredPeriod <> 0 Then
       freq = 1 / MeasuredPeriod
    Else
       freq = 0
    End If
End Function

Public Function Measure_Eye(ByVal Wave1 As DSPWave, ByRef RJ As Double, ByRef DDJ As Double, _
        ByRef MeasuredUI As Double, ByRef MidWidth As Double, ByRef LowWidth As Double, ByRef HighWidth As Double, _
        ByRef RiseTime As Double, ByRef FallTime As Double, ByRef dspStatus As Long, ByVal v As Double, ByRef EarlyLow As Double, ByRef LateLow As Double, ByRef EarlyHigh As Double, ByRef LateHigh As Double) As Long

    Dim EarlyMid As Double
    Dim LateMid As Double

    
    Call Wave1.MeasureEye(RJ, DDJ, MeasuredUI, RiseTime, FallTime, EarlyLow, LateLow, EarlyMid, LateMid, EarlyHigh, LateHigh, dspStatus)
    ''wave1.MeasureJitter Rj, DDj, MeasuredUI, dspStatus
    MidWidth = EarlyMid + MeasuredUI - LateMid
    LowWidth = EarlyLow + MeasuredUI - LateLow
    HighWidth = EarlyHigh + MeasuredUI - LateHigh

'    thehdw.Digital.Jitter.FileExport wave1, ".\AMPAPLLJitter_" & v & "V.txt"

    Debug.Print
End Function

Public Function Measure_Jitter( _
        CapDSPWave As DSPWave, _
        ByRef resultDspWave As DSPWave, _
        ByRef dspStatus As Long) As Long
        '
        Dim calcDspWave As New DSPWave
        Dim RJ As Double
        Dim DDJ As Double
        Dim MeasuredUI As Double
        
        'Create results DSPwave for more compact storage of test results
        Set resultDspWave = New DSPWave
        resultDspWave.CreateConstant 0, 3
        '
        calcDspWave = CapDSPWave.Copy
        'calcDspWave.Plot "Jitter Test:"
        Call calcDspWave.MeasureJitter(RJ, DDJ, MeasuredUI, dspStatus)
        '
        resultDspWave.Element(0) = RJ
        resultDspWave.Element(1) = DDJ
        resultDspWave.Element(2) = MeasuredUI

        
End Function

Public Function MeasureJitter_vb(ByVal wave As DSPWave, DDJ As Double, RJ As Double, Meas_UI As Double, dspStatus As Long, ByRef dutycycle As Double, ByRef freq As Double) As Long
    Dim PWHigh As Double
    Dim PWLow As Double
    Dim MeasuredPeriod As Double
   ' Dim dspStatus As Long
   ' Dim dspStatus1 As Long
    Dim UI As Double

    Call wave.SerialMeasureJitter(RJ, DDJ, Meas_UI, dspStatus)
    Call wave.measuretime(PWHigh, PWLow, MeasuredPeriod, dspStatus)
    dutycycle = PWHigh / MeasuredPeriod * 100
    If MeasuredPeriod <> 0 Then
       freq = 1 / MeasuredPeriod
    Else
       freq = 0
    End If
    End Function

Public Function analyzeEye_vb(ByVal wave As DSPWave, UI As Double, RJ As Double, DDJ As Double, RiseTime As Double, _
    FallTime As Double, EarlyLow As Double, LateLow As Double, EarlyMid As Double, LateMid As Double, EarlyHigh As Double, _
    LateHigh As Double, dspStatus As Long) As Long
    
    Call wave.SerialMeasureEye(RJ, DDJ, UI, RiseTime, FallTime, EarlyLow, LateLow, EarlyMid, LateMid, EarlyHigh, LateHigh, dspStatus)
End Function

Public Function Measuredutycycle(ByVal wave As DSPWave, dutycycle As Double, freq As Double) As Long
 
    Dim RJ        As Double
    Dim DDJ       As Double
    Dim UI        As Double
    Dim dspStatus As Long

    Call wave.SerialMeasureJitter(RJ, DDJ, UI, dspStatus)

    dutycycle = 100# * (UI - DDJ) / (2 * UI)
    freq = 1 / (2 * UI)
    
End Function



Public Function BitWf2Arry(ByVal InWf As DSPWave, ByVal WrdWdth As Integer, _
    ByRef NoOfSamples As Long, ByRef DataWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    Convert captured (serial) bit stream to data waveform, Assume LSB->MSB in the bit stream (reversed
    ''''        order may be easily accommodated by adding a switch in the argument list)
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''        BitWf2Arry is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        InWf          : DSP Wave (serial) to be converted
    ''''        WrdWdth  : number of bits per word
    ''''        NoOfSamples    : number of samples found in the bit stream
    ''''        DataWf         : converted (parallel) DSP Wave
    ''''
          
    NoOfSamples = InWf.SampleSize
    
    If NoOfSamples Mod WrdWdth <> 0 Then
         Debug.Print vbNewLine & "Bit stream wave size not integer times of the word width." _
            & " Waveform will Be truncated" & vbNewLine
    End If
    
    DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
    NoOfSamples = DataWf.SampleSize

End Function



Public Function DdrSwpWf2Arry(ByVal InWf As DSPWave, ByVal NoOfBists As Long, ByVal EyeStrobes As Long, _
    ByVal NoOfMdlls As Long, ByVal MdllWrdWidth As Long, ByRef EyeWf As DSPWave, ByRef MdllWf As DSPWave) As Long
    
    ''''--------------------------------------------------------------------------------------------------
    ''''    Convert captured DDR sweep bit stream waveform to data array, ssume LSB->MSB in the bit stream.
    ''''        order may be easily accommodated by adding a switch in the argument list)
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''        DdrSwpWf2Arry is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        InWf            : DSP Wave (serial) to be converted
    ''''        NoOfBists       : DDR LB test consists of individual blocks, suchas lanes, byte.
    ''''        EyeStrobes      : Number of strobes for each eye in this sweep
    ''''        NoOfMdlls       : Number of MDLL cal codes to be captured.
    ''''        MdllWrdWidth    : Number of bits in each MDLL cal code
    ''''        EyeWf           : Converted eye diagram waveform to be passed back
    ''''        MdllWf          : Converted MDLL code array in waveform to be passed back
        
    Dim NoOfSamples As Long
    Dim SegWf As New DSPWave, tmpwf As New DSPWave
    Dim BistIdx As Long, SegIdx   As Long
    Dim NoOfSegs As Long, SegWidth As Long
    Dim BistsPerMdll As Long    '''' MDLL is per lane, while BIST may be per byte.
                                '''' MDLL code will be captured per lane, followed by all the eyes in that lane
    
    '''' Not all loopback involve MDLL code
    If NoOfMdlls <> 0 Then
        If NoOfBists Mod NoOfMdlls <> 0 Then
            Debug.Print vbNewLine & "No of BISTs must be multiples of Number of MDLLs!" & vbNewLine
        End If
        NoOfSegs = NoOfMdlls
        BistsPerMdll = NoOfBists / NoOfMdlls
        SegWidth = BistsPerMdll * EyeStrobes + MdllWrdWidth
    Else
        BistsPerMdll = NoOfBists
        SegWidth = EyeStrobes * NoOfBists
        NoOfSegs = 1
    End If
        
    EyeWf.CreateConstant 0, EyeStrobes * NoOfBists, DspLong
    MdllWf.CreateConstant 0, NoOfMdlls, DspLong
      
    '''' construct the eye diagram and MDLL waveforms from the captured bit stream
    For SegIdx = 0 To NoOfSegs - 1
        SegWf = InWf.Select(SegIdx * SegWidth, 1, SegWidth).Copy
        If MdllWrdWidth <> 0 Then
            tmpwf = SegWf.Select(0, 1, MdllWrdWidth).Copy
            MdllWf.Element(SegIdx) = tmpwf.ConvertStreamTo(tldspParallel, MdllWrdWidth, 0, Bit0IsMsb).Element(0)
            tmpwf.Clear
        End If
        
        For BistIdx = 0 To BistsPerMdll - 1
            Dim st0 As Long, st1 As Long
            st0 = EyeStrobes * BistIdx
            st1 = SegIdx * (SegWidth - MdllWrdWidth) + st0
            tmpwf = SegWf.Select(MdllWrdWidth + st0, 1, EyeStrobes).Copy
            EyeWf.Select(st1, 1, EyeStrobes).Replace tmpwf
            tmpwf.Clear
        Next BistIdx
    Next SegIdx
    
End Function
Public Function FindMaxEyeWidth(ByVal SwqEyeWf As DSPWave, ByVal SwsEyeWf As DSPWave, _
                        ByVal NoOfBists As Integer, ByRef EyeWidthWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    Search for the eyewidth corresponding to the largest eye in the presense of one more eyes.
    ''''    An eye is defined as a continuous "1's"
    ''''    Based on the test methodologies used for TMA, there are 2 sweeps from the center of the UI, the first
    ''''        towards left, and second right, Swq and Sws respectively. The function would stitch them together by
    ''''        reversing the Swq waveform, and concatenating with the Sws wave.
    ''''
    ''''    The resulting eye diagram is from the left to right covering the entire UI.

    
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016

    ''''    Usage
    ''''        FindMaxEyeWidth is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        SwqEyeWf,SwsEyeWf   : Eye diagram waveforms, all BISTs concatenated.
    ''''        NoOfBists           : DDR LB test consists of individual blocks, suchas lanes, byte.
    ''''        EyeWidthWf          : Array of the eyewidths found, in DSP wave format, to be passed back
    
    ''''    rev 1, by Tim, 2/26/2016
    ''''
    ''''--------------------------------------------------------------------------------------------------
    ''''           New add to fix out of index error, exit for not an good method. Need to change it.
    ''''--------------------------------------------------------------------------------------------------

    Dim WholeEyeWf As New DSPWave
    Dim NoOfSamples As Long, EyeStrobes As Long
    Dim EyeWidth As Long
    Dim BistIdx As Long, idx As Long, idx2 As Long
    Dim MaxEye As Long
    EyeWidthWf.CreateConstant 0, NoOfBists, DspLong
    
    NoOfSamples = SwqEyeWf.SampleSize
    If NoOfSamples <> SwsEyeWf.SampleSize Then
             Debug.Print vbNewLine & "The lengths of Swq and Sws eye sweep waveform not consistent!" & vbNewLine
    End If
    
    EyeStrobes = NoOfSamples / NoOfBists
    
    For BistIdx = 0 To NoOfBists - 1
        WholeEyeWf = SwqEyeWf.Select(BistIdx * EyeStrobes, 1, EyeStrobes).Copy   '''' dummy operation, to allocate element
        
        '''' SWQ is from right to left, starting the middle. Reversing the eye diagram
        For idx = 0 To EyeStrobes - 1
            'idx2 = (BistIdx + 1) * EyeStrobes - 1 - idx
            WholeEyeWf.Element(idx) = SwqEyeWf.Element(idx)
        Next idx
        
        '''' stitch the reversed SWQ eye diagram to the SWK one
        WholeEyeWf = WholeEyeWf.Concatenate(SwsEyeWf.Select(BistIdx * EyeStrobes, 1, EyeStrobes))
        
        MaxEye = 0
        Dim EyeResultFlag As Boolean
        
        
        
        '''' finding the maximum eyewidth
        For idx = 0 To 2 * EyeStrobes - 1
            EyeResultFlag = False
            If WholeEyeWf.Element(idx) = 1 Then '''' starting a new eye
                EyeWidth = 0
                Do
                    EyeWidth = EyeWidth + 1
                    idx = idx + 1
               
                If idx >= 2 * EyeStrobes Then
                    EyeResultFlag = True
                ElseIf WholeEyeWf.Element(idx) = 0 Then
                    EyeResultFlag = True
                End If
               
               
               'Loop 'Until WholeEyeWf.Element(idx) = 0
                        
              '  If WholeEyeWf.Element(idx) = 0 Then EyeResultFlag = True
                
                
                'Loop Until idx >= 2 * EyeStrobes - 1 Or WholeEyeWf.Element(idx) = 0
                Loop Until EyeResultFlag = True
                
                    '''' >= workaround for not being able to escape the loop while maintain the wave index within the range
                    ''''    in the condition
                    ''' The next statement separately including the last strobe
                If idx = 2 * EyeStrobes - 1 Then EyeWidth = EyeWidth + WholeEyeWf.Element(idx)
                If MaxEye < EyeWidth Then MaxEye = EyeWidth
            End If
        Next idx
        
        EyeWidthWf.Element(BistIdx) = MaxEye
        
    Next BistIdx
    

End Function

Public Function FlexibleBitWf2Arry(ByVal InWf As DSPWave, ByVal StartIndex As Long, ByVal WrdWdth As Long, ByRef DataWf As DSPWave, ByRef DataWf_Binary As DSPWave) As Long
    
    'FlexibleBitWf2Arry
    ''''--------------------------------------------------------------------------------------------------
    ''''    Convert captured (serial) bit stream to data waveform, Assume LSB->MSB in the bit stream (reversed
    ''''        order may be easily accommodated by adding a switch in the argument list)
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''        BitWf2Arry is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        InWf          : DSP Wave (serial) to be converted
    ''''        WrdWdth  : number of bits per word
    ''''        NoOfSamples    : number of samples found in the bit stream
    ''''        DataWf         : converted (parallel) DSP Wave
    ''''

''    NoOfSamples = InWf.SampleSize

''    If NoOfSamples Mod WrdWdth <> 0 Then
''         Debug.Print vbNewLine & "Bit stream wave size not integer times of the word width." _
''            & " Waveform will Be truncated" & vbNewLine
''    End If

    DataWf_Binary = InWf.Select(StartIndex, , WrdWdth).Copy
    DataWf = InWf.Select(StartIndex, , WrdWdth).ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)

End Function
Public Function Split_Dspwave(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
        For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
End Function

Public Function CreatSerDigSrc(ByVal DataWf As DSPWave, ByVal BitWdthWf As DSPWave, ByVal SrcSize As Long, _
                        ByVal isIndDataRepeat As Boolean, ByVal isAllDataRepeat As Boolean, _
                        ByRef DigSrcWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    This function generates DSP waveform for DSSC digital source, to replace Create_DigSrc_Data,
    ''''        which was written in VBT module, not DSP. Other changes included remove all hard codings,
    ''''        increase the flexibility for various inputs and repetition combinations. The number of
    ''''        data (registers) are now flexible
    ''''
    ''''    This function is for serial source only. If parallel, the method of generating the waveform
    ''''        should be different for best efficiency
    ''''
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/25/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''        DataWf As DSPWave       Data to be sourced.  Each element (paralle, long) for each register
    ''''        BitWdthWf As DSPWave    Bit width info, should have the same number of element as in DataWf
    ''''        SrcSize As Long         If serial: total number of bits. If parallel, number of data
    ''''        isIndDataRepeat/isAllDataRepeat As Boolean  repeating each data/reg or repeat complete sequence
    ''''                                one must be true while the other false
    ''''        DigSrcWf As DSPWave     The generated waveform to be used for digital source
    
    Dim NoOfData As Long, NoOfRepeats As Long
    Dim DataIdx As Long
    Dim tmpwf As New DSPWave
    Dim SingleWordWf As New DSPWave         '''' hold one single data point
    
    '''' re-initiate the dig source wf
    Set DigSrcWf = New DSPWave
    
    ''''    Verifiy the number of samples in the data wf and bitwidth wf are the same
    NoOfData = DataWf.SampleSize
    If NoOfData <> BitWdthWf.SampleSize Then
        Debug.Print vbNewLine & "The sizes of data wf and width wf no matching!" & vbNewLine
    End If
    
    ''''    Calculate the number of repeats and check if integer
    NoOfRepeats = SrcSize \ BitWdthWf.CalcSum
    If NoOfRepeats * BitWdthWf.CalcSum <> SrcSize Then
        Debug.Print vbNewLine & "The total number of source bits not integer of the sum of the data widths!" & vbNewLine
    End If
    
    ''''
    '''' creat waveform based on the repeat scheme: per data repeat or repeat all
    '''' must be done one data at a time because the bitwdiths of the data may be different
    ''''
    If isIndDataRepeat Then
        For DataIdx = 0 To NoOfData - 1
            SingleWordWf = DataWf.Select(DataIdx, 1, 1).ConvertStreamTo( _
                tldspSerial, BitWdthWf.Element(DataIdx), 0, Bit0IsLsb)     '''' one data at a time. 1: stride. 1: size
            
            '''' concatenation only works on wafevorm with more than one element
            If DigSrcWf.SampleSize > 0 Then
                DigSrcWf = DigSrcWf.Concatenate(SingleWordWf.repeat(NoOfRepeats))
            Else
                DigSrcWf = SingleWordWf.repeat(NoOfRepeats)
            End If
        Next DataIdx
    ElseIf isAllDataRepeat Then
        For DataIdx = 0 To NoOfData - 1
            SingleWordWf = DataWf.Select(DataIdx, 1, 1).ConvertStreamTo( _
                tldspSerial, BitWdthWf.Element(DataIdx), 0, Bit0IsLsb)     '''' one data at a time. 1: stride. 1: size
            
            If DigSrcWf.SampleSize > 0 Then
                DigSrcWf = DigSrcWf.Concatenate(SingleWordWf)
            Else
                DigSrcWf = SingleWordWf
            End If
        Next DataIdx
        DigSrcWf = DigSrcWf.repeat(NoOfRepeats)
    End If
    
    
End Function
Public Function CreateFlexibleDSPWave(ByVal InWf As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long
    
    ''Dim SerialStream As New DSPWave
    ''Dim ParallelStream As New DSPWave
    ''
    ''ParallelStream.CreateConstant 17, 1, DspLong
    ''SerialStream = ParallelStream.ConvertStreamTo(tldspSerial, 12, 0, Bit0IsLsb)
    InWf = InWf.ConvertDataTypeTo(DspLong)
    DataWf = InWf.ConvertStreamTo(tldspSerial, WrdWdth, 0, Bit0IsLsb)
    
End Function
Public Function CreateFlexibleDSPWave_lpro(ByVal InWf As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave, ByRef fine_wf As DSPWave) As Long
    
    ''Dim SerialStream As New DSPWave
    ''Dim ParallelStream As New DSPWave
    ''
    ''ParallelStream.CreateConstant 17, 1, DspLong
    ''SerialStream = ParallelStream.ConvertStreamTo(tldspSerial, 12, 0, Bit0IsLsb)
    InWf = InWf.ConvertDataTypeTo(DspLong)
    DataWf = InWf.ConvertStreamTo(tldspSerial, WrdWdth, 0, Bit0IsMsb)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
    DataWf = fine_wf.Concatenate(DataWf)
    
End Function

Public Function SetupTrimCodeBit(ByVal InWf As DSPWave, ByVal b_SetupToBit0 As Boolean, ByVal BitIndex As Long, ByVal b_ControlNextBit As Boolean, ByRef DataWf As DSPWave) As Long
    
    DataWf = InWf
    If b_SetupToBit0 = True Then
        DataWf.Element(BitIndex) = 0
    Else
        DataWf.Element(BitIndex) = 1
    End If
    If b_ControlNextBit Then
        DataWf.Element(BitIndex - 1) = 1
    End If
End Function
Public Function ConvertToLongAndSerialToParrel(ByVal InWf As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long
    
    InWf = InWf.ConvertDataTypeTo(DspLong)
    
     If (InWf.SampleSize > 1) Then 'Check for Decimal...If binary convert to decimal or else copies the decimal directly to output waveform
        DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
    Else
        DataWf = InWf
    End If
   ' DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
End Function
Public Function CombineDSPWave_and_ConvertToLongAndSerialToParrel(ByVal FirstDSP As DSPWave, ByVal SecondDSP As DSPWave, _
                                                                  ByVal FirstLength As Long, ByVal SecondLength As Long, ByRef CombineDSP As DSPWave, _
                                                                  ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long
    

    CombineDSP.CreateConstant 0, FirstLength + SecondLength, DspLong
    FirstDSP = FirstDSP.ConvertDataTypeTo(DspLong)
    SecondDSP = SecondDSP.ConvertDataTypeTo(DspLong)
    CombineDSP.Select(0, 1, FirstLength).Replace (FirstDSP)
    CombineDSP.Select(FirstLength, 1, SecondLength).Replace (SecondDSP)
    
    CombineDSP = CombineDSP.ConvertDataTypeTo(DspLong)
    DataWf = CombineDSP.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
    
End Function
Public Function FindMaxEyeWidth_reverse(ByVal SwqEyeWf As DSPWave, ByVal SwsEyeWf As DSPWave, _
                        ByVal NoOfBists As Integer, ByRef EyeWidthWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    Search for the eyewidth corresponding to the largest eye in the presense of one more eyes.
    ''''    An eye is defined as a continuous "1's"
    ''''    Based on the test methodologies used for TMA, there are 2 sweeps from the center of the UI, the first
    ''''        towards left, and second right, Swq and Sws respectively. The function would stitch them together by
    ''''        reversing the Swq waveform, and concatenating with the Sws wave.
    ''''
    ''''    The resulting eye diagram is from the left to right covering the entire UI.

    
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016

    ''''    Usage
    ''''        FindMaxEyeWidth is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        SwqEyeWf,SwsEyeWf   : Eye diagram waveforms, all BISTs concatenated.
    ''''        NoOfBists           : DDR LB test consists of individual blocks, suchas lanes, byte.
    ''''        EyeWidthWf          : Array of the eyewidths found, in DSP wave format, to be passed back
    
    ''''    rev 1, by Tim, 2/26/2016
    ''''
    ''''--------------------------------------------------------------------------------------------------
    ''''           New add to fix out of index error, exit for not an good method. Need to change it.
    ''''--------------------------------------------------------------------------------------------------

    Dim WholeEyeWf As New DSPWave
    Dim NoOfSamples As Long, EyeStrobes As Long
    Dim EyeWidth As Long
    Dim BistIdx As Long, idx As Long, idx2 As Long
    Dim MaxEye As Long
    EyeWidthWf.CreateConstant 0, NoOfBists, DspLong
    
    NoOfSamples = SwqEyeWf.SampleSize
    If NoOfSamples <> SwsEyeWf.SampleSize Then
             Debug.Print vbNewLine & "The lengths of Swq and Sws eye sweep waveform not consistent!" & vbNewLine
    End If
    
    EyeStrobes = NoOfSamples / NoOfBists
    
    For BistIdx = 0 To NoOfBists - 1
        WholeEyeWf = SwqEyeWf.Select(BistIdx * EyeStrobes, 1, EyeStrobes).Copy   '''' dummy operation, to allocate element
        
        '''' SWQ is from right to left, starting the middle. Reversing the eye diagram
        For idx = 0 To EyeStrobes - 1
            idx2 = (BistIdx + 1) * EyeStrobes - 1 - idx
            WholeEyeWf.Element(idx) = SwqEyeWf.Element(idx2)
        Next idx
        
        '''' stitch the reversed SWQ eye diagram to the SWK one
        WholeEyeWf = WholeEyeWf.Concatenate(SwsEyeWf.Select(BistIdx * EyeStrobes, 1, EyeStrobes))
        
        MaxEye = 0
        Dim EyeResultFlag As Boolean
        
        
        
        '''' finding the maximum eyewidth
        For idx = 0 To 2 * EyeStrobes - 1
            EyeResultFlag = False
            If WholeEyeWf.Element(idx) = 1 Then '''' starting a new eye
                EyeWidth = 0
                Do
                    EyeWidth = EyeWidth + 1
                    idx = idx + 1
               
                If idx >= 2 * EyeStrobes Then
                    EyeResultFlag = True
                ElseIf WholeEyeWf.Element(idx) = 0 Then
                    EyeResultFlag = True
                End If
               
               
               'Loop 'Until WholeEyeWf.Element(idx) = 0
                        
              '  If WholeEyeWf.Element(idx) = 0 Then EyeResultFlag = True
                
                
                'Loop Until idx >= 2 * EyeStrobes - 1 Or WholeEyeWf.Element(idx) = 0
                Loop Until EyeResultFlag = True
                
                    '''' >= workaround for not being able to escape the loop while maintain the wave index within the range
                    ''''    in the condition
                    ''' The next statement separately including the last strobe
                If idx = 2 * EyeStrobes - 1 Then EyeWidth = EyeWidth + WholeEyeWf.Element(idx)
                If MaxEye < EyeWidth Then MaxEye = EyeWidth
            End If
        Next idx
        
        EyeWidthWf.Element(BistIdx) = MaxEye
        
    Next BistIdx
    

End Function

Public Function DSP_Add(ByRef InWf_1 As DSPWave, ByRef InWf_2 As DSPWave) As Long
    
    InWf_1 = InWf_1.Add(InWf_2)
    
End Function

Public Function DSP_Subtract(ByRef InWf_1 As DSPWave, ByRef InWf_2 As DSPWave) As Long
    
    InWf_1 = InWf_1.Subtract(InWf_2)
    
End Function

Public Function DSP_Multiply(ByRef InWf_1 As DSPWave, ByRef InWf_2 As DSPWave) As Long
    
    InWf_1 = InWf_1.Multiply(InWf_2)
    
End Function

Public Function DSP_Divide(ByRef InWf_1 As DSPWave, ByRef InWf_2 As DSPWave) As Long
    
    InWf_1 = InWf_1.Divide(InWf_2)
    
End Function

Public Function DSPWaveDecToBinary(ByVal InWf As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long
    InWf = InWf.ConvertDataTypeTo(DspLong)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
    DataWf = InWf.ConvertStreamTo(tldspSerial, WrdWdth, 0, Bit0IsMsb)
    
End Function

Public Function Transfer2GrayCode(ByVal InWf As DSPWave, ByRef OutWf As DSPWave, ByRef OutWf_Dec As DSPWave) As Long
''    Exit Function
    OutWf.CreateConstant 0, InWf.SampleSize, DspLong
    OutWf_Dec.CreateConstant 0, 1, DspLong
    
    Dim i As Long
    For i = 0 To InWf.SampleSize - 1 Step 1
        If i = 0 Then
            OutWf.Element(i) = InWf.Element(i)
        Else
            If InWf.Element(i - 1) = InWf.Element(i) Then
                OutWf.Element(i) = 0
            Else
                OutWf.Element(i) = 1
            End If
        End If
    Next i
    OutWf_Dec = OutWf.ConvertStreamTo(tldspParallel, OutWf.SampleSize, 0, Bit0IsMsb)
    
End Function

Public Function PreCheckMinMaxTrimCode(ByVal b_SetupToBit0 As Boolean, ByRef DataWf As DSPWave) As Long
    
    Dim i As Long
    For i = 0 To DataWf.SampleSize - 1
        If b_SetupToBit0 = True Then
            DataWf.Element(i) = 0
        Else
            DataWf.Element(i) = 1
        End If
    Next i
End Function

Public Function DSP_DivideConstant(ByRef InWf_1 As DSPWave, ByVal Denominator As Long) As Long
    If Denominator = 0 Then
''        TheExec.Datalog.WriteComment ("Error! Divide 0.")
        Exit Function
    Else
        InWf_1 = InWf_1.Divide(Denominator)
    End If
End Function


Public Function CombineDSPWave(ByVal FirstDSP As DSPWave, ByVal SecondDSP As DSPWave, ByVal FirstLength As Long, ByVal SecondLength As Long, ByRef CombineDSP As DSPWave) As Long
    Dim i As Long, j As Long
    Dim index As Long
    CombineDSP.CreateConstant 0, FirstLength + SecondLength, DspLong
    FirstDSP = FirstDSP.ConvertDataTypeTo(DspLong)
    SecondDSP = SecondDSP.ConvertDataTypeTo(DspLong)
    CombineDSP.Select(0, 1, FirstLength).Replace (FirstDSP)
    CombineDSP.Select(FirstLength, 1, SecondLength).Replace (SecondDSP)
End Function

Public Function SelectCertainBitsToDec(ByVal InWf As DSPWave, ByVal StartBit As Long, ByVal BitLength As Long, ByRef DataWf As DSPWave) As Long
    Dim TempDSP As New DSPWave
    InWf = InWf.ConvertDataTypeTo(DspLong)
    TempDSP = InWf.Select(StartBit, 1, BitLength).Copy
    
    DataWf = TempDSP.ConvertStreamTo(tldspParallel, BitLength, 0, Bit0IsMsb)
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
End Function
Public Function ConcatenateDSP(ByVal DSPWave_First As DSPWave, ByVal First_StartElement As Long, ByVal First_EndElement As Long, _
                               ByVal DSPWave_Second As DSPWave, ByVal Second_StartElement As Long, ByVal Second_EndElement As Long, _
                               ByRef DSPWave_Combine As DSPWave) As Long

    Dim FinalLength As Long
    Dim i As Long
    FinalLength = Abs(First_EndElement - First_StartElement) + Abs(Second_EndElement - Second_StartElement) + 2
    DSPWave_Combine.CreateConstant 0, FinalLength
    
    Dim b_MinToMax_First As Boolean
    Dim b_MinToMax_Second As Boolean
    Dim Step_First As Integer
    Dim Step_Second As Integer
    Dim counter As Long
    counter = 0
    If First_EndElement - First_StartElement > 0 Then
        b_MinToMax_First = True
        Step_First = 1
    Else
        b_MinToMax_First = False
        Step_First = -1
    End If
    
    If Second_EndElement - Second_StartElement > 0 Then
        b_MinToMax_Second = True
        Step_Second = 1
    Else
        b_MinToMax_Second = False
        Step_Second = -1
    End If
    
    For i = First_StartElement To First_EndElement Step Step_First
        DSPWave_Combine.Element(counter) = DSPWave_First.Element(i)
        counter = counter + 1
    Next i


    For i = Second_StartElement To Second_EndElement Step Step_Second
        DSPWave_Combine.Element(counter) = DSPWave_Second.Element(i)
        counter = counter + 1
    Next i
End Function

Public Function DSP_BitWiseAnd(ByVal InputDSP As DSPWave, ByVal FixedDSP As DSPWave, ByVal BitWidth As Long, ByRef OutputDSP As DSPWave) As Long
    InputDSP = InputDSP.ConvertDataTypeTo(DspLong)
    OutputDSP.CreateConstant 0, BitWidth, DspLong
    OutputDSP = InputDSP.bitwiseand(FixedDSP)
End Function


Public Function DSP_BitWiseOr(ByVal InputDSP As DSPWave, ByVal FixedDSP As DSPWave, ByVal BitWidth As Long, ByRef OutputDSP As DSPWave) As Long
    InputDSP = InputDSP.ConvertDataTypeTo(DspLong)
    OutputDSP.CreateConstant 0, BitWidth, DspLong
    OutputDSP = InputDSP.BitwiseOr(FixedDSP)
End Function

Public Function DSP_BitWiseXOR(ByVal InputDSP As DSPWave, ByVal FixedDSP As DSPWave, ByVal BitWidth As Long, ByRef OutputDSP As DSPWave) As Long
    InputDSP = InputDSP.ConvertDataTypeTo(DspLong)
    OutputDSP.CreateConstant 0, BitWidth, DspLong
    OutputDSP = InputDSP.BitwiseXor(FixedDSP)
End Function


Public Function BinToDec(ByVal InWf As DSPWave, ByRef DataWf As DSPWave) As Long
    Dim WrdWdth As Long
    InWf = InWf.ConvertDataTypeTo(DspLong)
    WrdWdth = InWf.SampleSize
    DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)

End Function

Public Function DSP_Convert_2S_Complement(ByVal InWf As DSPWave, WordWidth As Long, DataWf As DSPWave) As Long
    InWf = InWf.ConvertDataTypeTo(DspLong)
    InWf.WordWidth = WordWidth
    Set DataWf = InWf.ConvertDataTypeTo(DspLong)
    Debug.Print DataWf.Element(0)
    
End Function

Sub WordWidthExample()
Dim FourBitValues As New DSPWave
Dim IntegerValues As New DSPWave
Dim i As Long
FourBitValues.CreateRamp 0, 1, 16, DspLong
FourBitValues.WordWidth = 4
Set IntegerValues = FourBitValues.ConvertDataTypeTo(DspLong)
For i = 0 To 15
    Debug.Print FourBitValues.Element(i); " => "; IntegerValues.Element(i)
Next
End Sub
Public Function DSP_ConvertDataTypeToLong(ByRef InWf_1 As DSPWave) As Long
    InWf_1.ConvertDataTypeTo (DspLong)
End Function

Public Function DSP_GrayCode2Bin(ByVal IsUnsigned As Boolean, ByVal InWf As DSPWave, ByRef OutWf As DSPWave, ByRef OutWf_Dec As DSPWave) As Long

    OutWf.CreateConstant 0, InWf.SampleSize, DspLong
    OutWf_Dec.CreateConstant 0, 1, DspLong
    
    Dim i As Long
    Dim MSB_ElementNumForSignUnsign As Long
    MSB_ElementNumForSignUnsign = InWf.SampleSize - 1
    
    Dim SignUnsignDiffBit As Long
    If IsUnsigned Then
        SignUnsignDiffBit = 1
    Else
        SignUnsignDiffBit = 2
    End If
    Dim index As Long
    index = 0
    For i = InWf.SampleSize - SignUnsignDiffBit To 0 Step -1
        If index = 0 Then
            OutWf.Element(i) = InWf.Element(i)
        Else
            If InWf.Element(i) = OutWf.Element(i + 1) Then
                OutWf.Element(i) = 0
            Else
                OutWf.Element(i) = 1
            End If
        End If
        index = index + 1
    Next i
    OutWf_Dec = OutWf.ConvertStreamTo(tldspParallel, OutWf.SampleSize, 0, Bit0IsMsb)
    
    If IsUnsigned = True Then
    
    Else
        If InWf.Element(MSB_ElementNumForSignUnsign) = 1 Then
            OutWf_Dec.Element(0) = OutWf_Dec.Element(0) * -1
        Else
        End If
    End If
    
End Function

Public Function DSP_2S_Complement_To_SignDec(InWf As DSPWave, WordWidth As Long, DataWf_DEC As DSPWave) As Long
''    Exit Function
    
    InWf = InWf.ConvertDataTypeTo(DspLong)
''''    InWf.WordWidth = WordWidth
''    Set DataWf = InWf.ConvertNumFormatTo(SignMagnitude, WordWidth)
''    Debug.Print DataWf.Element(0)
    Dim MaxVal As Long
    MaxVal = 2 ^ (WordWidth - 1)
    Dim DSPWaveWithoutSignBit As New DSPWave
    Dim SignBit As Long
    DSPWaveWithoutSignBit = InWf.Select(0, , WordWidth - 1).Copy
    SignBit = InWf.Element(WordWidth - 1)
    Dim DSPWaveWithoutSignBit_DEC As New DSPWave
    DSPWaveWithoutSignBit_DEC.CreateConstant 0, 1, DspLong
''    Call rundsp.BinToDec(DSPWaveWithoutSignBit, DSPWaveWithoutSignBit_DEC)
    Dim WrdWdthWithoutSignBit As Long
    ''InWf = InWf.ConvertDataTypeTo(DspLong)
    WrdWdthWithoutSignBit = DSPWaveWithoutSignBit.SampleSize
    DSPWaveWithoutSignBit_DEC = DSPWaveWithoutSignBit.ConvertStreamTo(tldspParallel, WrdWdthWithoutSignBit, 0, Bit0IsMsb)
    
    If SignBit = 0 Then
        DataWf_DEC = DSPWaveWithoutSignBit_DEC
    Else
        DataWf_DEC = DSPWaveWithoutSignBit_DEC.Subtract(MaxVal)
    End If
End Function

Public Function DSP_SignedBin_To_SignDec(InWf As DSPWave, WordWidth As Long, DataWf_DEC As DSPWave) As Long
''    Exit Function
    
    InWf = InWf.ConvertDataTypeTo(DspLong)
''''    InWf.WordWidth = WordWidth
''    Set DataWf = InWf.ConvertNumFormatTo(SignMagnitude, WordWidth)
''    Debug.Print DataWf.Element(0)
    Dim MaxVal As Long
    MaxVal = 2 ^ (WordWidth - 1)
    Dim DSPWaveWithoutSignBit As New DSPWave
    Dim SignBit As Long
    DSPWaveWithoutSignBit = InWf.Select(0, , WordWidth - 1).Copy
    SignBit = InWf.Element(WordWidth - 1)
    Dim DSPWaveWithoutSignBit_DEC As New DSPWave
    DSPWaveWithoutSignBit_DEC.CreateConstant 0, 1, DspLong
''    Call rundsp.BinToDec(DSPWaveWithoutSignBit, DSPWaveWithoutSignBit_DEC)
    Dim WrdWdthWithoutSignBit As Long
    ''InWf = InWf.ConvertDataTypeTo(DspLong)
    WrdWdthWithoutSignBit = DSPWaveWithoutSignBit.SampleSize
    DSPWaveWithoutSignBit_DEC = DSPWaveWithoutSignBit.ConvertStreamTo(tldspParallel, WrdWdthWithoutSignBit, 0, Bit0IsMsb)
    
    If SignBit = 0 Then
        DataWf_DEC = DSPWaveWithoutSignBit_DEC
    Else
        DataWf_DEC = DSPWaveWithoutSignBit_DEC.Negate
    End If
End Function

Public Function LPDPRX_EyeSweep(ByVal InputDSPWave As DSPWave, ByVal FinalEyeOutBitNum As Long, ByRef CalcOutputDSPWave As DSPWave, ByRef CalcEyeWidth As Long) As Long
    Dim i As Long
    Dim index As Long
    Dim MaxWitth As Long
    Dim TempMaxWidth As Long
    Dim FinalMaxWitth As Long
    
    CalcOutputDSPWave.CreateConstant 0, FinalEyeOutBitNum, DspLong
    
    For i = 0 To InputDSPWave.SampleSize - 1 Step 2
        If InputDSPWave.Element(i) = 0 And InputDSPWave.Element(i + 1) = 32768 Then
            CalcOutputDSPWave.Element(index) = 0
        Else
            CalcOutputDSPWave.Element(index) = 1
        End If
        index = index + 1
    Next i
'    CalcOutputDSPWave.Element(1) = 0
'    CalcOutputDSPWave.Element(2) = 0
'    CalcOutputDSPWave.Element(10) = 0
'    CalcOutputDSPWave.Element(11) = 0
'    CalcOutputDSPWave.Element(12) = 0
'    CalcOutputDSPWave.Element(21) = 0
'    CalcOutputDSPWave.Element(22) = 0
    MaxWitth = 0
    Dim All_zero_Flag As Boolean
    All_zero_Flag = True
    
    For i = 0 To CalcOutputDSPWave.SampleSize - 1
        If CalcOutputDSPWave.Element(i) = 0 Then
            MaxWitth = MaxWitth + 1
            TempMaxWidth = MaxWitth
        Else
            All_zero_Flag = False
            If FinalMaxWitth < TempMaxWidth Then
                FinalMaxWitth = TempMaxWidth
            End If
            MaxWitth = 0
        End If
    Next i
    If All_zero_Flag = True Then
        FinalMaxWitth = TempMaxWidth
    End If
    
    CalcEyeWidth = FinalMaxWitth
End Function

Public Function PCIE_EyeSweep(ByVal InputDSPWave As DSPWave, ByVal FinalEyeOutBitNum As Long, ByRef CalcOutputDSPWave As DSPWave, ByRef CalcEyeWidth As Long) As Long
    Dim i As Long
    Dim index As Long
    Dim MaxWitth As Long
    Dim TempMaxWidth As Long
    Dim FinalMaxWitth As Long
    
    CalcOutputDSPWave.CreateConstant 0, FinalEyeOutBitNum, DspLong
    
    For i = 0 To InputDSPWave.SampleSize - 1 Step 2
        If InputDSPWave.Element(i) = 0 And InputDSPWave.Element(i + 1) = 0 Then
            CalcOutputDSPWave.Element(index) = 0
        Else
            CalcOutputDSPWave.Element(index) = 1
        End If
        index = index + 1
    Next i
'    CalcOutputDSPWave.Element(1) = 0
'    CalcOutputDSPWave.Element(2) = 0
'    CalcOutputDSPWave.Element(10) = 0
'    CalcOutputDSPWave.Element(11) = 0
'    CalcOutputDSPWave.Element(12) = 0
'    CalcOutputDSPWave.Element(21) = 0
'    CalcOutputDSPWave.Element(22) = 0
    MaxWitth = 0
    Dim All_zero_Flag As Boolean
    All_zero_Flag = True
    
    For i = 0 To CalcOutputDSPWave.SampleSize - 1
        If CalcOutputDSPWave.Element(i) = 0 Then
            MaxWitth = MaxWitth + 1
            TempMaxWidth = MaxWitth
        Else
            All_zero_Flag = False
            If FinalMaxWitth < TempMaxWidth Then
                FinalMaxWitth = TempMaxWidth
            End If
            MaxWitth = 0
        End If
    Next i
    If All_zero_Flag = True Then
        FinalMaxWitth = TempMaxWidth
    End If
    
    CalcEyeWidth = FinalMaxWitth
End Function


Public Function SeprateDSP(DSP_Input As DSPWave, DSP_Input_UpperBIN As DSPWave, DSP_Input_BelowBIN As DSPWave) As Long
    Dim length As Long
    length = DSP_Input.SampleSize / 2
    DSP_Input_UpperBIN = DSP_Input.Select(0, , length).Copy
    DSP_Input_BelowBIN = DSP_Input.Select(0 + length, , length).Copy
End Function

Public Function AveWithStdev(InputWave As DSPWave, mean As Double, Std As Double) As Long
    mean = InputWave.CalcMeanWithStdDev(Std)
End Function


Public Function SeprateDSP_TTR_Single(DSP_Input_Update As DSPWave, DSP_Input_UpperBIN As DSPWave, DSP_Input_BelowBIN As DSPWave, ByRef DSP_Input_UpperDEC As DSPWave, ByRef DSP_Input_BelowDEC As DSPWave) As Long
 
    Dim length As Long
    Dim i As Long
    Dim WrdWdth As Long
    length = 8
    
    For i = 0 To DSP_Input_Update.SampleSize - 1 Step 16
        If i = 0 Then
            DSP_Input_UpperBIN = DSP_Input_Update.Select(i, , length).Copy
            DSP_Input_BelowBIN = DSP_Input_Update.Select(i + length, , length).Copy
        Else
            DSP_Input_UpperBIN = DSP_Input_UpperBIN.Concatenate(DSP_Input_Update.Select(i, , length))
            DSP_Input_BelowBIN = DSP_Input_BelowBIN.Concatenate(DSP_Input_Update.Select(i + length, , length))
        End If
    Next i
            
    DSP_Input_UpperDEC = DSP_Input_UpperBIN.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_BelowDEC = DSP_Input_BelowBIN.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    
End Function
Public Function SeprateDSP_TTR(DSP_Input_Update As DSPWave, DSP_Input_UpperBIN_1 As DSPWave, DSP_Input_UpperBIN_2 As DSPWave, DSP_Input_BelowBIN_1 As DSPWave, DSP_Input_BelowBIN_2 As DSPWave, ByRef DSP_Input_UpperDEC_1 As DSPWave, ByRef DSP_Input_UpperDEC_2 As DSPWave, ByRef DSP_Input_BelowDEC_1 As DSPWave, ByRef DSP_Input_BelowDEC_2 As DSPWave) As Long
 
    Dim length As Long
    Dim i As Long
    Dim WrdWdth As Long
    length = 8

    For i = 0 To DSP_Input_Update.SampleSize - 1 Step 32
        If i = 0 Then
            DSP_Input_UpperBIN_1 = DSP_Input_Update.Select(i, , length).Copy
            DSP_Input_UpperBIN_2 = DSP_Input_Update.Select(i + length, , length).Copy
            DSP_Input_BelowBIN_1 = DSP_Input_Update.Select(i + length * 2, , length).Copy
            DSP_Input_BelowBIN_2 = DSP_Input_Update.Select(i + length * 3, , length).Copy
        Else
            DSP_Input_UpperBIN_1 = DSP_Input_UpperBIN_1.Concatenate(DSP_Input_Update.Select(i, , length))
            DSP_Input_UpperBIN_2 = DSP_Input_UpperBIN_2.Concatenate(DSP_Input_Update.Select(i + length, , length))
            DSP_Input_BelowBIN_1 = DSP_Input_BelowBIN_1.Concatenate(DSP_Input_Update.Select(i + length * 2, , length))
            DSP_Input_BelowBIN_2 = DSP_Input_BelowBIN_2.Concatenate(DSP_Input_Update.Select(i + length * 3, , length))
        
        End If
    Next i
            
    DSP_Input_UpperDEC_1 = DSP_Input_UpperBIN_1.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_UpperDEC_2 = DSP_Input_UpperBIN_2.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_BelowDEC_1 = DSP_Input_BelowBIN_1.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    DSP_Input_BelowDEC_2 = DSP_Input_BelowBIN_2.ConvertStreamTo(tldspParallel, length, 0, Bit0IsMsb)
    
End Function
Public Function FindMaxEyeWidth_reverse_bywidth(ByVal SwqEyeWf As DSPWave, ByVal SwsEyeWf As DSPWave, _
                        ByVal Cont_width As DSPWave, ByRef EyeWidthWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    Search for the eyewidth corresponding to the largest eye in the presense of one more eyes.
    ''''    An eye is defined as a continuous "1's"
    ''''    Based on the test methodologies used for TMA, there are 2 sweeps from the center of the UI, the first
    ''''        towards left, and second right, Swq and Sws respectively. The function would stitch them together by
    ''''        reversing the Swq waveform, and concatenating with the Sws wave.
    ''''
    ''''    The resulting eye diagram is from the left to right covering the entire UI.

    
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016

    ''''    Usage
    ''''        FindMaxEyeWidth is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        SwqEyeWf,SwsEyeWf   : Eye diagram waveforms, all BISTs concatenated.
    ''''        NoOfBists           : DDR LB test consists of individual blocks, suchas lanes, byte.
    ''''        EyeWidthWf          : Array of the eyewidths found, in DSP wave format, to be passed back
    
    ''''    rev 1, by Tim, 2/26/2016
    ''''
    ''''--------------------------------------------------------------------------------------------------
    ''''           New add to fix out of index error, exit for not an good method. Need to change it.
    ''''--------------------------------------------------------------------------------------------------

    Dim WholeEyeWf As New DSPWave
    Dim NoOfSamples As Long, EyeStrobes As Long
    Dim EyeWidth As Long
    Dim BistIdx As Long, idx As Long, idx2 As Long
    Dim MaxEye As Long
    Dim Lane_count As Long
    Dim EyeStrobes_split As Long
    Dim Count_Eyerunbit As Long
    
    Lane_count = Cont_width.SampleSize
    
    
    NoOfSamples = SwqEyeWf.SampleSize
    
    If NoOfSamples <> SwsEyeWf.SampleSize Then
             Debug.Print vbNewLine & "The lengths of Swq and Sws eye sweep waveform not consistent!" & vbNewLine
        
    End If
    
   EyeWidthWf.CreateConstant 0, Lane_count, DspLong
    
    For BistIdx = 0 To Lane_count - 1
    
       EyeStrobes = Cont_width.Element(BistIdx)
       
            If BistIdx = 0 Then
                
                  EyeStrobes_split = 0
                  
                  Count_Eyerunbit = Cont_width.Element(BistIdx)
                
            Else
                
                  EyeStrobes_split = EyeStrobes_split + Cont_width.Element(BistIdx - 1)
                  
                  Count_Eyerunbit = Count_Eyerunbit + Cont_width.Element(BistIdx)
'                  Count_Eyerunbit = Count_Eyerunbit + Cont_width.Element(BistIdx - 1)
            End If
        
        WholeEyeWf = SwqEyeWf.Select(EyeStrobes_split, 1, EyeStrobes).Copy    '''' dummy operation, to allocate element
        
        '''' SWQ is from right to left, starting the middle. Reversing the eye diagram
        For idx = 0 To EyeStrobes - 1
            idx2 = Count_Eyerunbit - 1 - idx
            WholeEyeWf.Element(idx) = SwqEyeWf.Element(idx2)
        Next idx
        
        '''' stitch the reversed SWQ eye diagram to the SWK one
        WholeEyeWf = WholeEyeWf.Concatenate(SwsEyeWf.Select(EyeStrobes_split, 1, EyeStrobes))
        
        MaxEye = 0
        Dim EyeResultFlag As Boolean
        
        
        
        '''' finding the maximum eyewidth
        For idx = 0 To 2 * EyeStrobes - 1
            EyeResultFlag = False
            If WholeEyeWf.Element(idx) = 1 Then '''' starting a new eye
                EyeWidth = 0
                Do
                    EyeWidth = EyeWidth + 1
                    idx = idx + 1
               
                If idx >= 2 * EyeStrobes Then
                    EyeResultFlag = True
                ElseIf WholeEyeWf.Element(idx) = 0 Then
                    EyeResultFlag = True
                End If
               
               
               'Loop 'Until WholeEyeWf.Element(idx) = 0
                        
              '  If WholeEyeWf.Element(idx) = 0 Then EyeResultFlag = True
                
                
                'Loop Until idx >= 2 * EyeStrobes - 1 Or WholeEyeWf.Element(idx) = 0
                Loop Until EyeResultFlag = True
                
                    '''' >= workaround for not being able to escape the loop while maintain the wave index within the range
                    ''''    in the condition
                    ''' The next statement separately including the last strobe
                If idx = 2 * EyeStrobes - 1 Then EyeWidth = EyeWidth + WholeEyeWf.Element(idx)
                If MaxEye < EyeWidth Then MaxEye = EyeWidth
            End If
        Next idx
        
        EyeWidthWf.Element(BistIdx) = MaxEye
        
    Next BistIdx
End Function

Public Function DspWaveMergeRepeat(ByRef OutputDspWave As DSPWave, ByVal InDSPwave As DSPWave, ByVal SampleSize As Long) As Long
    
    Dim InDspWaveSampleSize As Long
    
    InDspWaveSampleSize = InDSPwave.SampleSize
    OutputDspWave = InDSPwave.repeat(CLng(SampleSize / InDspWaveSampleSize))
End Function

Public Function DSPWf_Concatenate(OutputDspWave As DSPWave, InDSPwave As DSPWave, dummy As Long) As Long

    If OutputDspWave.SampleSize = 0 Then
        OutputDspWave = InDSPwave.Copy
    Else
        OutputDspWave = OutputDspWave.ConvertDataTypeTo(DspLong)
        InDSPwave = InDSPwave.ConvertDataTypeTo(DspLong)
        OutputDspWave = OutputDspWave.Concatenate(InDSPwave)
    End If

End Function

Public Function MTR_ASGMTR_Freq_Calculation(ByRef InWf1 As DSPWave, ByRef InWf2 As DSPWave, ByRef InWf3 As DSPWave, ByRef InWf4 As DSPWave, ByRef InWf5 As DSPWave, ByRef InWf6 As DSPWave, ByRef InWf7 As DSPWave, ByRef InWf8 As DSPWave, ByRef InWf9 As DSPWave, ByRef InWf10 As DSPWave, ByRef InWf11 As DSPWave, ByRef InWf12 As DSPWave, ByRef InWf13 As DSPWave, ByRef InWf14 As DSPWave, ByRef InWf15 As DSPWave, ByRef InWf16 As DSPWave, ByRef InWf17 As DSPWave, ByRef InWf18 As DSPWave, ByRef InWf19 As DSPWave, ByRef InWf20 As DSPWave, ByRef InWf21 As DSPWave, ByRef InWf22 As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long

Dim DataWf1 As New DSPWave: DataWf1 = InWf1.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf2 As New DSPWave: DataWf2 = InWf2.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf3 As New DSPWave: DataWf3 = InWf3.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf4 As New DSPWave: DataWf4 = InWf4.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf5 As New DSPWave: DataWf5 = InWf5.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf6 As New DSPWave: DataWf6 = InWf6.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf7 As New DSPWave: DataWf7 = InWf7.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf8 As New DSPWave: DataWf8 = InWf8.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf9 As New DSPWave: DataWf9 = InWf9.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf10 As New DSPWave: DataWf10 = InWf10.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf11 As New DSPWave: DataWf11 = InWf11.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf12 As New DSPWave: DataWf12 = InWf12.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf13 As New DSPWave: DataWf13 = InWf13.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf14 As New DSPWave: DataWf14 = InWf14.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf15 As New DSPWave: DataWf15 = InWf15.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf16 As New DSPWave: DataWf16 = InWf16.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf17 As New DSPWave: DataWf17 = InWf17.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf18 As New DSPWave: DataWf18 = InWf18.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf19 As New DSPWave: DataWf19 = InWf19.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf20 As New DSPWave: DataWf20 = InWf20.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf21 As New DSPWave: DataWf21 = InWf21.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)
Dim DataWf22 As New DSPWave: DataWf22 = InWf22.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(50000)

InWf1 = InWf1.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf2 = InWf2.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf3 = InWf3.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf4 = InWf4.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf5 = InWf5.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf6 = InWf6.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf7 = InWf7.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf8 = InWf8.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf9 = InWf9.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf10 = InWf10.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf11 = InWf11.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf12 = InWf12.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf13 = InWf13.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf14 = InWf14.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf15 = InWf15.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf16 = InWf16.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf17 = InWf17.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf18 = InWf18.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf19 = InWf19.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf20 = InWf20.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf21 = InWf21.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf22 = InWf22.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)

DataWf = DataWf1.Concatenate(DataWf2).Concatenate(DataWf3).Concatenate(DataWf4).Concatenate(DataWf5).Concatenate(DataWf6).Concatenate(DataWf7).Concatenate(DataWf8).Concatenate(DataWf9).Concatenate(DataWf10).Concatenate(DataWf11).Concatenate(DataWf12).Concatenate(DataWf13).Concatenate(DataWf14).Concatenate(DataWf15).Concatenate(DataWf16).Concatenate(DataWf17).Concatenate(DataWf18).Concatenate(DataWf19).Concatenate(DataWf20).Concatenate(DataWf21).Concatenate(DataWf22)


End Function

Public Function MTR_DSGMTR_Freq_Calculation(ByRef InWf1 As DSPWave, ByRef InWf2 As DSPWave, ByRef InWf3 As DSPWave, ByRef InWf4 As DSPWave, ByRef InWf5 As DSPWave, ByRef InWf6 As DSPWave, ByRef InWf7 As DSPWave, ByRef InWf8 As DSPWave, ByRef InWf9 As DSPWave, ByRef InWf10 As DSPWave, ByRef InWf11 As DSPWave, ByRef InWf12 As DSPWave, ByRef InWf13 As DSPWave, ByRef InWf14 As DSPWave, ByRef InWf15 As DSPWave, ByRef InWf16 As DSPWave, ByRef InWf17 As DSPWave, ByRef InWf18 As DSPWave, ByRef InWf19 As DSPWave, ByRef InWf20 As DSPWave, ByRef InWf21 As DSPWave, ByRef InWf22 As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave) As Long

Dim DataWf1 As New DSPWave: DataWf1 = InWf1.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf2 As New DSPWave: DataWf2 = InWf2.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf3 As New DSPWave: DataWf3 = InWf3.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf4 As New DSPWave: DataWf4 = InWf4.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf5 As New DSPWave: DataWf5 = InWf5.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf6 As New DSPWave: DataWf6 = InWf6.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf7 As New DSPWave: DataWf7 = InWf7.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf8 As New DSPWave: DataWf8 = InWf8.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf9 As New DSPWave: DataWf9 = InWf9.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf10 As New DSPWave: DataWf10 = InWf10.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf11 As New DSPWave: DataWf11 = InWf11.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf12 As New DSPWave: DataWf12 = InWf12.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf13 As New DSPWave: DataWf13 = InWf13.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf14 As New DSPWave: DataWf14 = InWf14.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf15 As New DSPWave: DataWf15 = InWf15.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf16 As New DSPWave: DataWf16 = InWf16.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf17 As New DSPWave: DataWf17 = InWf17.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf18 As New DSPWave: DataWf18 = InWf18.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf19 As New DSPWave: DataWf19 = InWf19.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf20 As New DSPWave: DataWf20 = InWf20.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf21 As New DSPWave: DataWf21 = InWf21.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)
Dim DataWf22 As New DSPWave: DataWf22 = InWf22.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb).Multiply(93750)

InWf1 = InWf1.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf2 = InWf2.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf3 = InWf3.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf4 = InWf4.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf5 = InWf5.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf6 = InWf6.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf7 = InWf7.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf8 = InWf8.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf9 = InWf9.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf10 = InWf10.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf11 = InWf11.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf12 = InWf12.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf13 = InWf13.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf14 = InWf14.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf15 = InWf15.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf16 = InWf16.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf17 = InWf17.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf18 = InWf18.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf19 = InWf19.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf20 = InWf20.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf21 = InWf21.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
InWf22 = InWf22.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)

DataWf = DataWf1.Concatenate(DataWf2).Concatenate(DataWf3).Concatenate(DataWf4).Concatenate(DataWf5).Concatenate(DataWf6).Concatenate(DataWf7).Concatenate(DataWf8).Concatenate(DataWf9).Concatenate(DataWf10).Concatenate(DataWf11).Concatenate(DataWf12).Concatenate(DataWf13).Concatenate(DataWf14).Concatenate(DataWf15).Concatenate(DataWf16).Concatenate(DataWf17).Concatenate(DataWf18).Concatenate(DataWf19).Concatenate(DataWf20).Concatenate(DataWf21).Concatenate(DataWf22)


End Function

Public Function Split_Dspwave_PCIETXPLL(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, calc_data As DSPWave, delta_value As DSPWave, target_var As Double, ByVal calibration_target_value As Long, ByVal start_search As Long, BinTarget As DSPWave, delta_value2 As DSPWave, target_var2 As Double, BinTarget2 As DSPWave) As Long
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).Copy
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    '=====================================================20180904
    
    'Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp1_dict As New DSPWave
    Dim temp2_dict As New DSPWave
    Dim temp_delta_value As Integer
    
    
    ''''calc and print in datalog
'    temp1_dict.CreateConstant 0, 1, DspLong
'    temp2_dict.CreateConstant 0, 1, DspLong
'
'    For i = 0 To CLng((OutWf.SampleSize) / 2 - 1)
'            temp1_dict.Element(0) = OutWf.Element(i)
'            temp2_dict.Element(0) = OutWf.Element(i + 32)
'            calc_data.Element(i) = (temp2_dict.Element(0) + temp1_dict.Element(0)) / 2
'    Next i
    For i = start_search To CLng(OutWf.SampleSize) - 1
        calc_data.Element(i - start_search) = OutWf.Element(i)
    Next i
    ''' compare the target

    temp_delta_value = 9999
                 
    For k = start_search To CLng((OutWf.SampleSize)) - 1
        delta_value.Element(k) = Abs(calibration_target_value - calc_data.Element(k - start_search))
        'search min delta
        If delta_value.Element(k) < temp_delta_value Then
            temp_delta_value = delta_value.Element(k)
            target_var = k
        End If
    Next k
    
    ''''''''''''' Decimal to Binary dspwave for dictionary'''''''''''''''
    
    Dim TempVal As Long
    
    TempVal = target_var
    For i = 0 To CLng((BinTarget.SampleSize)) - 1
        BinTarget.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            BinTarget.Element(i) = 0
        End If
    Next i

        
    '=====================================================20180904
    
    
End Function

Public Function Split_Dspwave_CIOCALC(Outwf_T1 As DSPWave, V2 As Double, V3 As Double, V4 As Double, V5 As Double, V6 As Double, V7 As Double, V8 As Double, V9 As Double, V10 As Double, V11 As Double, V12 As Double, V13 As Double, V14 As Double, V15 As Double, V16 As Double, V17 As Double, V18 As Double, V19 As Double, V20 As Double, V21 As Double, V22 As Double, V23 As Double, _
                                V24 As Double, V25 As Double, storeDSP As DSPWave) As Long
                                                                                                                                                                                                                                                               
                                                                                                                                                                                                                                                               
    ' Special calculation for T1
    ' 1.1 store measured data
    Dim d_temp(24) As Double
    d_temp(0) = V2
    d_temp(1) = V3
    d_temp(2) = V4
    d_temp(3) = V5
    d_temp(4) = V6
    d_temp(5) = V7
    d_temp(6) = V8
    d_temp(7) = V9
    d_temp(8) = V10
    d_temp(9) = V11
    d_temp(10) = V12
    d_temp(11) = V13
    d_temp(12) = V14
    d_temp(13) = V15
    d_temp(14) = V16
    d_temp(15) = V17
    d_temp(16) = V18
    d_temp(17) = V19
    d_temp(18) = V20
    d_temp(19) = V21
    d_temp(20) = V22
    d_temp(21) = V23
    d_temp(22) = V24
    d_temp(23) = V25
    ' 1.2 Calculate desired index
    Dim target_index As Long: target_index = 999
    Dim target_gap As Double: target_gap = 2013144
    Dim i As Long
    For i = 0 To 23
        If d_temp(i) >= 0.3 And (Outwf_T1.Element(i) - 3333) >= 0 Then
            target_gap = Abs(Outwf_T1.Element(i) - 3333)
            target_index = i
            Exit For  'once larger then use it for 1st read
        End If
    Next
    
    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_gap = 2013144
        For i = 0 To 23
            If d_temp(i) >= 0.3 And Abs(Outwf_T1.Element(i) - 3333) < target_gap Then
                target_gap = Abs(Outwf_T1.Element(i) - 3333)
                target_index = i
            End If
        Next
    End If
'
    Outwf_T1.Element(24) = target_index
    
    Dim TempVal As Long
    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    Outwf_T1.Element(24) = TempVal
    
    For i = 0 To CLng((storeDSP.SampleSize)) - 1
        storeDSP.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            storeDSP.Element(i) = 0
        End If
    Next i
                                                                                                                                                                                                                                                               
End Function

Public Function Split_Dspwave_CIOPLL(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, calc_data As DSPWave, delta_value As DSPWave, target_var_low As Double, target_var_high As Double, target_var_low2 As Double, target_var_high2 As Double, ByVal calibration_target_value_low As Long, ByVal calibration_target_value_high As Long, ByVal start_search As Long, Bin_Target_low As DSPWave, Bin_Target_high As DSPWave, Bin_Target_low2 As DSPWave, Bin_Target_high2 As DSPWave) As Long
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).Copy
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    '=====================================================20180904
    
    'Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp1_dict As New DSPWave
    Dim temp2_dict As New DSPWave
    Dim temp_delta_value As Integer
    
    
    ''''calc and print in datalog
'    temp1_dict.CreateConstant 0, 1, DspLong
'    temp2_dict.CreateConstant 0, 1, DspLong
'
'    For i = 0 To CLng((OutWf.SampleSize) / 2 - 1)
'            temp1_dict.Element(0) = OutWf.Element(i)
'            temp2_dict.Element(0) = OutWf.Element(i + 32)
'            calc_data.Element(i) = (temp2_dict.Element(0) + temp1_dict.Element(0)) / 2
'    Next i
    For i = start_search To CLng(OutWf.SampleSize) - 1
        calc_data.Element(i - start_search) = OutWf.Element(i)
    Next i
    ''' compare the target

    temp_delta_value = 9999
                 
    For k = start_search To CLng((OutWf.SampleSize)) - 1
        delta_value.Element(k) = Abs(calibration_target_value_low - OutWf.Element(k))
        'search min delta
        If delta_value.Element(k) < temp_delta_value Then
            temp_delta_value = delta_value.Element(k)
            target_var_low = k
        End If
    Next k
    
    temp_delta_value = 9999
    
    For k = start_search To CLng((OutWf.SampleSize)) - 1
        delta_value.Element(k) = Abs(calibration_target_value_high - OutWf.Element(k))
        'search min delta
        If delta_value.Element(k) < temp_delta_value Then
            temp_delta_value = delta_value.Element(k)
            target_var_high = k
        End If
    Next k
    '=====================================================20180904
    Dim TempVal As Long

    
    TempVal = target_var_low
    For i = 0 To CLng((Bin_Target_low.SampleSize)) - 1
        Bin_Target_low.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            Bin_Target_low.Element(i) = 0
        End If
    Next i

    TempVal = target_var_high
    For i = 0 To CLng((Bin_Target_low.SampleSize)) - 1
        Bin_Target_high.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            Bin_Target_high.Element(i) = 0
        End If
    Next i
    
End Function

Public Function Split_Dspwave_AUS(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, Outwf_T1 As DSPWave, sda_measuredata3 As Double, sda_measuredata4 As Double, sda_measuredata5 As Double, sda_measuredata6 As Double, sda_measuredata7 As Double, sda_measuredata8 As Double, sda_measuredata9 As Double, sda_measuredata10 As Double, sda_measuredata11 As Double, sda_measuredata12 As Double, sda_measuredata13 As Double, sda_measuredata14 As Double, sda_measuredata15 As Double, sda_measuredata16 As Double, sda_measuredata17 As Double, sda_measuredata18 As Double, sda_measuredata19 As Double, sda_measuredata20 As Double, sda_measuredata21 As Double, sda_measuredata22 As Double, sda_measuredata23 As Double, sda_measuredata24 As Double, sda_measuredata25 As Double, _
                                sda_measuredata26 As Double, storeDSP As DSPWave) As Long
                                                                                                                                                                                                                                                               
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim temp_dsp As New DSPWave
    Dim Count As Long
    current_loc = 0
    Count = 0
                                                                                                                                                                                                                                                                '' 20170705 - Solve for over 32 bits to decimal
    temp_dsp.CreateConstant 0, width_Wf.SampleSize, DspDouble
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).Copy
        If width_Wf.Element(i) = 16 Then
            'dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
            temp_dsp.Element(Count) = dec_wave.Element(0)
            Count = Count + 1
            'current_loc = current_loc + width_Wf.Element(i)
        End If
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    Outwf_T1 = temp_dsp.Select(1, 1, 25).Copy
                                                                                                                                                                                                                                                               
    ' Special calculation for T1
    ' 1.1 store measured data
    Dim d_temp(24) As Double
    d_temp(0) = sda_measuredata3
    d_temp(1) = sda_measuredata4
    d_temp(2) = sda_measuredata5
    d_temp(3) = sda_measuredata6
    d_temp(4) = sda_measuredata7
    d_temp(5) = sda_measuredata8
    d_temp(6) = sda_measuredata9
    d_temp(7) = sda_measuredata10
    d_temp(8) = sda_measuredata11
    d_temp(9) = sda_measuredata12
    d_temp(10) = sda_measuredata13
    d_temp(11) = sda_measuredata14
    d_temp(12) = sda_measuredata15
    d_temp(13) = sda_measuredata16
    d_temp(14) = sda_measuredata17
    d_temp(15) = sda_measuredata18
    d_temp(16) = sda_measuredata19
    d_temp(17) = sda_measuredata20
    d_temp(18) = sda_measuredata21
    d_temp(19) = sda_measuredata22
    d_temp(20) = sda_measuredata23
    d_temp(21) = sda_measuredata24
    d_temp(22) = sda_measuredata25
    d_temp(23) = sda_measuredata26
    ' 1.2 Calculate desired index
    Dim target_index As Long: target_index = 999
    Dim target_gap As Double: target_gap = 2013144
    For i = 0 To 23
        If d_temp(i) >= 0.3 And (Outwf_T1.Element(i) - 3333) >= 0 Then
            target_gap = Abs(Outwf_T1.Element(i) - 3333)
            target_index = i
            Exit For  'once larger then use it for 1st read
        End If
    Next
    
    If target_index = 999 Then ' if  there  is no value above 3333 then use closest one
        target_gap = 2013144
        For i = 0 To 23
            If d_temp(i) >= 0.3 And Abs(Outwf_T1.Element(i) - 3333) < target_gap Then
                target_gap = Abs(Outwf_T1.Element(i) - 3333)
                target_index = i
            End If
        Next
    End If
'
    Outwf_T1.Element(24) = target_index
    
    Dim TempVal As Long
    If target_index > 15 Then
        TempVal = target_index + 8
    Else
        TempVal = target_index
    End If
    Outwf_T1.Element(24) = TempVal
    
    For i = 0 To CLng((storeDSP.SampleSize)) - 1
        storeDSP.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 5 Then
            storeDSP.Element(i) = 0
        End If
    Next i
                                                                                                                                                                                                                                                               
End Function

Public Function Split_Dspwave_CIO(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, Outwf_T1 As DSPWave, Outwf_T2 As DSPWave) As Long
    
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim temp_dsp As New DSPWave
    Dim Count As Long
    current_loc = 0
    Count = 0
    
    temp_dsp.CreateConstant 0, width_Wf.SampleSize, DspDouble
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).Copy
        If width_Wf.Element(i) = 16 Then
            temp_dsp.Element(Count) = dec_wave.Element(0)
            Count = Count + 1
        End If

        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    Outwf_T1 = temp_dsp.Select(1, 1, Count / 2 - 2).Copy
    Outwf_T2 = temp_dsp.Select(Count / 2 + 1, 1, Count / 2 - 2).Copy
    
End Function

Public Function Split_Dspwave_PCIEREFPLL(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave, calc_data As DSPWave, delta_value As DSPWave, target_var As Double, ByVal calibration_target_value As Long, ByVal start_search As Long, BinTarget As DSPWave) As Long
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
    For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb).Copy
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
    
    '=====================================================20180904
    
    'Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp1_dict As New DSPWave
    Dim temp2_dict As New DSPWave
    Dim temp_delta_value As Integer
    
    
    ''''calc and print in datalog
'    temp1_dict.CreateConstant 0, 1, DspLong
'    temp2_dict.CreateConstant 0, 1, DspLong
'
'    For i = 0 To CLng((OutWf.SampleSize) / 2 - 1)
'            temp1_dict.Element(0) = OutWf.Element(i)
'            temp2_dict.Element(0) = OutWf.Element(i + 32)
'            calc_data.Element(i) = (temp2_dict.Element(0) + temp1_dict.Element(0)) / 2
'    Next i
    For i = start_search To CLng(OutWf.SampleSize) - 1
        calc_data.Element(i - start_search) = OutWf.Element(i)
    Next i
    ''' compare the target

    temp_delta_value = 9999
                 
    For k = start_search To CLng((OutWf.SampleSize)) - 1
        delta_value.Element(k) = Abs(calibration_target_value - calc_data.Element(k - start_search))
        'search min delta
        If delta_value.Element(k) < temp_delta_value Then
            temp_delta_value = delta_value.Element(k)
            target_var = k
        End If
    Next k
    
    ''''''''''''' Decimal to Binary dspwave for dictionary'''''''''''''''
    
    Dim TempVal As Long
    
    TempVal = target_var
    For i = 0 To CLng((BinTarget.SampleSize)) - 1
        BinTarget.Element(i) = TempVal Mod 2
        TempVal = TempVal \ 2
        If i >= 6 Then
            BinTarget.Element(i) = 0
        End If
    Next i

        
    '=====================================================20180904
    
    
End Function

Public Function Split_2SComplementDSPWave_To_SignDec(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    Dim DSP_2sComplement As New DSPWave
    current_loc = 0
    For i = 0 To width_Wf.SampleSize - 1
        DSP_2sComplement = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        Call DSP_2S_Complement_To_SignDec(DSP_2sComplement, width_Wf.Element(i), dec_wave)
        If i = 0 Then
            OutWf = dec_wave
        Else
            OutWf = OutWf.Concatenate(dec_wave)
        End If
        current_loc = current_loc + width_Wf.Element(i)
    Next i
End Function

Public Function Split_Gray_2sComplementDSPWave_to_Dec(DSPSignedGray_StartBit As DSPWave, DSPUnSignedGray_StartBit As DSPWave, DSP2sComplement_StartBit As DSPWave, DSPSignedBin_StartBit As DSPWave, InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
Dim dec_wave As New DSPWave, current_loc As Long, i As Long
Dim Index_SignedGray As Long: Index_SignedGray = 0
Dim Index_UnSignedGray As Long: Index_UnSignedGray = 0
Dim Index_2sComplement As Long: Index_2sComplement = 0
Dim Index_SignedBin As Long: Index_SignedBin = 0
Dim DSP_SignedGray As New DSPWave
Dim DSP_UnSignedGray As New DSPWave
Dim DSP_2sComplement As New DSPWave
Dim DSP_SignedBin As New DSPWave
Dim DSP_UnSignedBinary As New DSPWave
Dim Out_Wf_Binary As New DSPWave
current_loc = 0
For i = 0 To width_Wf.SampleSize - 1
    If current_loc = DSPSignedGray_StartBit.Element(Index_SignedGray) Then
        DSP_SignedGray = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        Call DSP_GrayCode2Bin(False, DSP_SignedGray, Out_Wf_Binary, dec_wave)
        If Index_SignedGray <> DSPSignedGray_StartBit.SampleSize - 1 Then
            Index_SignedGray = Index_SignedGray + 1
        End If
    ElseIf current_loc = DSPUnSignedGray_StartBit.Element(Index_UnSignedGray) Then
        DSP_UnSignedGray = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        Call DSP_GrayCode2Bin(True, DSP_UnSignedGray, Out_Wf_Binary, dec_wave)
        If Index_UnSignedGray <> DSPUnSignedGray_StartBit.SampleSize - 1 Then
            Index_UnSignedGray = Index_UnSignedGray + 1
        End If
    ElseIf current_loc = DSP2sComplement_StartBit.Element(Index_2sComplement) Then
        DSP_2sComplement = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        Call DSP_2S_Complement_To_SignDec(DSP_2sComplement, width_Wf.Element(i), dec_wave)
        If Index_2sComplement <> DSP2sComplement_StartBit.SampleSize - 1 Then
            Index_2sComplement = Index_2sComplement + 1
        End If
    ElseIf current_loc = DSPSignedBin_StartBit.Element(Index_SignedBin) Then
        DSP_SignedBin = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        Call DSP_SignedBin_To_SignDec(DSP_SignedBin, width_Wf.Element(i), dec_wave)
        If Index_SignedBin <> DSPSignedBin_StartBit.SampleSize - 1 Then
            Index_SignedBin = Index_SignedBin + 1
        End If
    Else
        DSP_UnSignedBinary = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        dec_wave = DSP_UnSignedBinary.ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
    End If
    If i = 0 Then
        OutWf = dec_wave
    Else
        OutWf = OutWf.Concatenate(dec_wave)
    End If
    current_loc = current_loc + width_Wf.Element(i)
Next i
End Function

Public Function DSP_Opt_EYE(ByRef DqSwpWf As DSPWave, ByVal DSP_Eye_StartBit_DQ As DSPWave, ByVal DSP_Eye_BitLength_DQ As DSPWave, ByRef DqsSwpWf As DSPWave, ByVal DSP_Eye_StartBit_DQS As DSPWave, ByVal DSP_Eye_BitLength_DQS As DSPWave, ByVal NoOfBists As Integer, ByRef DQ_EYE_Data As DSPWave, ByRef DQS_EYE_Data As DSPWave, ByRef DSP_Eye_Width As DSPWave) As Long

    Dim i As Integer

    For i = 0 To DSP_Eye_StartBit_DQ.SampleSize - 1
        If i = 0 Then
            DQ_EYE_Data = DqSwpWf.Select(DSP_Eye_StartBit_DQ.Element(i), , DSP_Eye_BitLength_DQ.Element(i)).Copy
        Else
             DQ_EYE_Data = DQ_EYE_Data.Concatenate(DqSwpWf.Select(DSP_Eye_StartBit_DQ.Element(i), , DSP_Eye_BitLength_DQ.Element(i)))
        End If
    Next i

    For i = 0 To DSP_Eye_StartBit_DQS.SampleSize - 1
        If i = 0 Then
            DQS_EYE_Data = DqsSwpWf.Select(DSP_Eye_StartBit_DQS.Element(i), , DSP_Eye_BitLength_DQS.Element(i)).Copy
        Else
            DQS_EYE_Data = DQS_EYE_Data.Concatenate(DqsSwpWf.Select(DSP_Eye_StartBit_DQS.Element(i), , DSP_Eye_BitLength_DQS.Element(i)))
        End If
    Next i

    Call FindMaxEyeWidth_reverse(DQ_EYE_Data, DQS_EYE_Data, NoOfBists, DSP_Eye_Width)
    
End Function

Public Function Split_Gray_to_Dec(DSPSignedGray_StartBit As DSPWave, DSPUnSignedGray_StartBit As DSPWave, InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
Dim dec_wave As New DSPWave, current_loc As Long, i As Long
Dim Index_SignedGray As Long: Index_SignedGray = 0
Dim Index_UnSignedGray As Long: Index_UnSignedGray = 0
Dim Index_2sComplement As Long: Index_2sComplement = 0
Dim DSP_SignedGray As New DSPWave
Dim DSP_UnSignedGray As New DSPWave
Dim DSP_2sComplement As New DSPWave
Dim DSP_UnSignedBinary As New DSPWave
Dim Out_Wf_Binary As New DSPWave
current_loc = 0
For i = 0 To width_Wf.SampleSize - 1
    If current_loc = DSPSignedGray_StartBit.Element(Index_SignedGray) Then
        DSP_SignedGray = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        Call DSP_GrayCode2Bin(False, DSP_SignedGray, Out_Wf_Binary, dec_wave)
        If Index_SignedGray <> DSPSignedGray_StartBit.SampleSize - 1 Then
            Index_SignedGray = Index_SignedGray + 1
        End If
    ElseIf current_loc = DSPUnSignedGray_StartBit.Element(Index_UnSignedGray) Then
        DSP_UnSignedGray = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        Call DSP_GrayCode2Bin(True, DSP_UnSignedGray, Out_Wf_Binary, dec_wave)
        If Index_UnSignedGray <> DSPUnSignedGray_StartBit.SampleSize - 1 Then
            Index_UnSignedGray = Index_UnSignedGray + 1
        End If
'    ElseIf current_loc = DSP2sComplement_StartBit.Element(Index_2sComplement) Then
'        DSP_2sComplement = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
'        Call DSP_2S_Complement_To_SignDec(DSP_2sComplement, width_Wf.Element(i), dec_wave)
'        If Index_2sComplement <> DSP2sComplement_StartBit.SampleSize - 1 Then
'            Index_2sComplement = Index_2sComplement + 1
'        End If
    Else
        DSP_UnSignedBinary = InWf.Select(current_loc, , width_Wf.Element(i)).Copy
        dec_wave = DSP_UnSignedBinary.ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsMsb)
    End If
    If i = 0 Then
        OutWf = dec_wave
    Else
        OutWf = OutWf.Concatenate(dec_wave)
    End If
    current_loc = current_loc + width_Wf.Element(i)
Next i
End Function

''20190604AddFunction
Public Function DSPWf_Dec2Binary(ByVal InWf As DSPWave, ByVal DataWdth As Long, ByRef OutWf As DSPWave) As Long
    
    OutWf = InWf.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, DataWdth, 0, Bit0IsMsb)
    
End Function
Public Function ReAssignmentDSPWave(ByVal InWf As DSPWave, ByVal SeparateNum As Long, ByRef OutWf As DSPWave, ByVal Src_dig As Boolean, ByVal assignment As DSPWave, ByVal AssignmentDSPWave As DSPWave) As Long
    Dim i As Integer    'Dylan Edited 20190528
    Dim j As Integer
    Dim RegSize As Long
    Dim InWfSize As Long
    Dim FullSize As Long
    
    
    InWfSize = InWf.SampleSize
    FullSize = OutWf.SampleSize
    RegSize = FullSize / SeparateNum
    
    If Src_dig = False Then
        For i = 1 To SeparateNum
            For j = 1 To RegSize
                OutWf.Element((i * RegSize) - j) = InWf.Element(RegSize - j)
                'Debug.Print i * RegSize - j
            Next j
        Next i
    Else
        For i = 1 To SeparateNum
            If AssignmentDSPWave.Element(i - 1) = 0 Then
               For j = 1 To RegSize
                OutWf.Element((i * RegSize) - j) = InWf.Element(RegSize - j)
                'Debug.Print i * RegSize - j
               Next j
            Else
               For j = 1 To RegSize
               OutWf.Element((i * RegSize) - j) = assignment.Element(RegSize - j)
               Next j
            End If
        Next
      
    End If
      
End Function

Public Function ElementTransformer(ByRef InWf As DSPWave, ByVal SeparateNum As Long, ByVal ElementOffset As Long) As Long
    Dim i As Long       'Dylan Edited 20190528
    Dim SizeCnt As Long
    Dim ProcessDSP As New DSPWave
    SizeCnt = InWf.SampleSize
    ProcessDSP.CreateConstant 0, SizeCnt, DspLong
    
    If ElementOffset = 0 Then
        ElementOffset = ElementOffset - 1
        For i = 0 To SizeCnt - 1
            ProcessDSP.Element(i) = InWf.Element(ElementOffset - i)
        Next i
        
        For i = 0 To SizeCnt - 1
            InWf.Element(ElementOffset - SeparateNum + i) = ProcessDSP.Element(i)
        Next i
    Else
        For i = 0 To SizeCnt - 1
            ProcessDSP.Element(i) = InWf.Element(SizeCnt - i - 1)
        Next i
        
        For i = 0 To SizeCnt - 1
            InWf.Element(i) = ProcessDSP.Element(i)
        Next i
    End If
    
    
    
    
    
    
    
'''''    j = SeparateNum / 2
'''''    SeparateNum = SeparateNum - 1                           ' Minimum element is zero
'''''    ElementOffset = ElementOffset - 1                       ' MAxmum element is Maxmum - 1
'''''    ProcessDSP.CreateConstant 0, SeparateNum, DspLong
'''''
'''''
'''''    For i = 0 To j - 1
'''''        ProcessDSP.Element(i) = InWf.Element(ElementOffset - i)
'''''        InWf.Element(ElementOffset - i) = InWf.Element(ElementOffset - SeparateNum + i)
'''''        InWf.Element(ElementOffset - SeparateNum + i) = ProcessDSP.Element(i)
'''''    Next i

End Function
Public Function SetupLinearTrimCodeBit(ByVal TrimMethod As Boolean, ByRef TrimCode As Double, ByVal b_SetupToBit0 As Boolean, _
ByVal RegSize As Long, ByRef DataWf As DSPWave, ByVal doallFlag As Boolean) As Long
' Dylan Edited 20190529
    Dim i As Integer
    Dim SizeTemp As Integer
    Dim TrimTemp As Integer
    Dim TotallySize As Integer
    
    SizeTemp = CInt(RegSize) - 1
    TotallySize = CInt(DataWf.SampleSize) - 1
    
    If TrimMethod = True Then                                   ' Increase Linear
        If doallFlag = True Then
            TrimCode = TrimCode + 1
        ElseIf b_SetupToBit0 = False Then
            TrimCode = TrimCode + 1
        End If
    Else                                                        ' Decrease Linear
        If doallFlag = True Then
            TrimCode = TrimCode - 1
        ElseIf b_SetupToBit0 = False Then
            TrimCode = TrimCode - 1
        End If
    End If
    
    TrimTemp = TrimCode
    If b_SetupToBit0 = False Or doallFlag = True Then
        For i = 0 To SizeTemp
            If TrimTemp <> 0 Then
                If TrimTemp \ (2 ^ (SizeTemp - i)) <> 0 Then
                    DataWf.Element(TotallySize - i) = 1
                Else
                    DataWf.Element(TotallySize - i) = 0
                End If
                If i <> SizeTemp Then
                    If TrimTemp >= (2 ^ (SizeTemp - i)) Then
                        TrimTemp = TrimTemp - (2 ^ (SizeTemp - i))
                    End If
                End If
            Else
                DataWf.Element(TotallySize - i) = 0
            End If
        Next i
    End If
    
End Function

Public Function CalculateDSPWaveforTrimCode(ByVal InWf As DSPWave, ByVal WrdWdth As Long, ByRef DataWf As DSPWave, _
ByVal TrimTotallyOffset As Integer, ByVal TrimBasedNum As Integer, ByRef FinallyWf As DSPWave) As Long
' Dylan Edited 20190529
    Dim i As Integer
    Dim CalculateDSPWave As New DSPWave
    
    If InWf.SampleSize <> 1 Then                                                                    ' Avoid sweep fail which any site
        CalculateDSPWave.CreateConstant 0, 1, DspLong
        InWf = InWf.ConvertDataTypeTo(DspLong)
        DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsMsb)
        DataWf.Element(0) = TrimBasedNum - (TrimTotallyOffset - DataWf.Element(0))
        
        
        If DataWf.Element(0) >= 0 Then                                                              ' Transfer to Binary if element(0) is positive
            CalculateDSPWave.Element(0) = DataWf.Element(0)
            FinallyWf = CalculateDSPWave.ConvertStreamTo(tldspSerial, WrdWdth, 0, Bit0IsMsb)        ' Transfer from Decimal to Binary
        Else                                                                                        ' Transfer to 2'S if element(0) is negative
            CalculateDSPWave.Element(0) = Abs(DataWf.Element(0))
            If CalculateDSPWave.Element(0) > (2 ^ (WrdWdth - 1) - 1) Then                           ' Illegal jugment
                FinallyWf.CreateConstant 0, 1, DspLong
            Else
                CalculateDSPWave.Element(0) = DataWf.Element(0)
                CalculateDSPWave.Element(0) = CalculateDSPWave.Element(0) + 2 ^ WrdWdth             ' Calculate 2'S complement
                FinallyWf = CalculateDSPWave.ConvertStreamTo(tldspSerial, WrdWdth, 0, Bit0IsMsb)    ' Transfer from Binary to Decimal
            End If
        End If
    End If
End Function

Public Function SetupBinaryTrimCodeBit(ByVal InWf As DSPWave, ByVal b_SetupToBit0 As Boolean, ByVal BitIndex As Long, ByVal InitStateByCapCode As Long, _
ByVal TrimOffset As Long, ByVal TrimOriginalSize As Long, ByRef DataWf As DSPWave, ByVal b_ControlNextBit As Boolean, ByVal AssignmentDSPWave As DSPWave, _
ByVal CoverSize As Long) As Long
'Dylan Edited 20190615
       
    Dim i As Long
    
    Dim Square As Long
    Dim SelsetOffset As Long
    Dim CalculateSize As Long
    Dim CalculateDSP As New DSPWave
    Dim assigment_select As Long
    CalculateSize = CLng(DataWf.SampleSize)
    DataWf = InWf

    For i = 0 To AssignmentDSPWave.SampleSize
        If AssignmentDSPWave.Element(i) = 0 Then
            SelsetOffset = i * CalculateSize
            Exit For
        Else
            SelsetOffset = 0
        End If
    Next i


    CalculateDSP = DataWf.Select(SelsetOffset, 1, CalculateSize)
    DataWf = CalculateDSP.Copy
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
    DataWf = DataWf.ConvertStreamTo(tldspParallel, CalculateSize, 0, Bit0IsMsb)
    If DataWf.Element(0) <> 0 Then
        DataWf.Element(0) = DataWf.Element(0) + 1
        DataWf.Element(0) = DataWf.Element(0) - TrimOffset
    End If
    Square = (2) ^ CInt(TrimOriginalSize)
    
    
    If InitStateByCapCode = 1 Then
        If b_SetupToBit0 = True Then
            DataWf.Element(0) = DataWf.Element(0) - (Square / ((2) ^ (CInt(TrimOriginalSize) - BitIndex)))
        Else
            DataWf.Element(0) = DataWf.Element(0) + (Square / ((2) ^ (CInt(TrimOriginalSize) - BitIndex)))
        End If
    Else
        If b_SetupToBit0 = True Then
            DataWf.Element(0) = DataWf.Element(0) + (Square / ((2) ^ (CInt(TrimOriginalSize) - BitIndex)))
        Else
            DataWf.Element(0) = DataWf.Element(0) - (Square / ((2) ^ (CInt(TrimOriginalSize) - BitIndex)))
        End If
    End If

    DataWf.Element(0) = DataWf.Element(0) - 1
    DataWf.Element(0) = DataWf.Element(0) + TrimOffset
    CalculateDSP = DataWf.Copy
    CalculateDSP = CalculateDSP.ConvertDataTypeTo(DspLong)
'    DataWf = CalculateDSP.ConvertStreamTo(tldspSerial, TrimOriginalSize + 1, 0, Bit0IsMsb)
    DataWf = CalculateDSP.ConvertStreamTo(tldspSerial, CoverSize, 0, Bit0IsMsb)
    
    
    
''    CalculateSize = DataWf.SampleSize
''    CalculateDSP.CreateConstant 0, 1, DspLong
''    DataWf = DataWf.ConvertDataTypeTo(DspLong)
''    CalculateDSP = DataWf.ConvertStreamTo(tldspParallel, CalculateSize, 0, Bit0IsMsb)
''    CalculateDSP.Element(0) = CalculateDSP.Element(0) - TrimOffset      ' Retrieve original trim code
''    CalculateDSP = CalculateDSP.ConvertDataTypeTo(DspLong)
''    CalculateDSP = CalculateDSP.ConvertStreamTo(tldspSerial, TrimOriginalSize, 0, Bit0IsMsb)

'    If InitStateByCapCode = 0 Then                  '                   |-----------------|-----------------|
'        If b_SetupToBit0 = True Then                ' Code              0                128               256
'            CalculateDSP.Element(BitIndex) = 0      ' Distributed       0000000000000011111111111111111111111
'        Else
'            CalculateDSP.Element(BitIndex) = 1
'        End If
'
'
'    Else
'        If b_SetupToBit0 = True Then                '                   |-----------------|-----------------|
'            CalculateDSP.Element(BitIndex) = 1      ' Code              0                128               256
'        Else                                        ' Distributed       1111111111111000000000000000000000000
'            CalculateDSP.Element(BitIndex) = 0
'        End If
'    End If
'    CalculateDSP = CalculateDSP.ConvertDataTypeTo(DspLong)
'    CalculateDSP = CalculateDSP.ConvertStreamTo(tldspParallel, TrimOriginalSize, 0, Bit0IsMsb)
'    CalculateDSP.Element(0) = CalculateDSP.Element(0) + TrimOffset      ' Addition trimoffset value
'    CalculateDSP = CalculateDSP.ConvertDataTypeTo(DspLong)
'    DataWf = CalculateDSP.ConvertStreamTo(tldspSerial, CalculateSize, 0, Bit0IsMsb)
    
    
End Function

'Public Function DSPWf_Dec2Binary(ByVal InWf As DSPWave, ByVal DataWdth As Long, ByRef OutWf As DSPWave) As Long
'
'    OutWf = InWf.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, DataWdth, 0, Bit0IsMsb)
'
'End Function

Public Function BitWf2Arry_MSB1st(ByVal InWf As DSPWave, ByVal WrdWdth As Integer, _
    ByRef NoOfSamples As Long, ByRef DataWf As DSPWave) As Long
    ''''--------------------------------------------------------------------------------------------------
    ''''    Convert captured (serial) bit stream to data waveform, Assume MSB->LSB in the bit stream (reversed
    ''''        order may be easily accommodated by adding a switch in the argument list)
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''        BitWf2Arry is to be called by a VBT function
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''
    ''''        InWf          : DSP Wave (serial) to be converted
    ''''        WrdWdth  : number of bits per word
    ''''        NoOfSamples    : number of samples found in the bit stream
    ''''        DataWf         : converted (parallel) DSP Wave
    ''''
          
    NoOfSamples = InWf.SampleSize
    
    If NoOfSamples Mod WrdWdth <> 0 Then
         Debug.Print vbNewLine & "Bit stream wave size not integer times of the word width." _
            & " Waveform will Be truncated" & vbNewLine
    End If
    
    DataWf = InWf.ConvertStreamTo(tldspParallel, WrdWdth, 0, Bit0IsLsb)
    NoOfSamples = DataWf.SampleSize

End Function

Public Function Split_Dspwave_MSB1st(ByVal InWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
    'Dim Split_Wave_ary(2) As New DSPWave
    Dim dec_wave As New DSPWave, current_loc As Long, i As Long
    current_loc = 0
    '' 20170705 - Solve for over 32 bits to decimal
    OutWf.CreateConstant 0, width_Wf.SampleSize, DspDouble
        For i = 0 To width_Wf.SampleSize - 1
        dec_wave = InWf.Select(current_loc, , width_Wf.Element(i)).ConvertStreamTo(tldspParallel, width_Wf.Element(i), 0, Bit0IsLsb)
        OutWf.Element(i) = dec_wave.Element(0)
        current_loc = current_loc + width_Wf.Element(i)
    Next i
End Function
Public Function SetupLinearTrimCodeBit_Linear(ByVal TrimMethod As Boolean, ByRef TrimCode As Double, ByVal b_SetupToBit0 As Boolean, _
ByVal RegSize As Long, ByRef DataWf As DSPWave, ByVal doallFlag As Boolean) As Long
' Dylan Edited 20190529
       Dim i As Integer
    Dim TempDSPWave As New DSPWave
    Dim SizeTemp As Integer
    Dim TrimTemp As Integer
    Dim TotallySize As Integer
    
    SizeTemp = CInt(RegSize) - 1
    TotallySize = CInt(DataWf.SampleSize) - 1
    
    If TrimMethod = True Then                                   ' Increase Linear
        If doallFlag = True Then
            TrimCode = TrimCode + 1
        ElseIf b_SetupToBit0 = False Then
            TrimCode = TrimCode + 1
        End If
    Else                                                        ' Decrease Linear
        If doallFlag = True Then
            TrimCode = TrimCode - 1
        ElseIf b_SetupToBit0 = False Then
            TrimCode = TrimCode - 1
        End If
    End If
    
    
    DataWf = DataWf.ConvertDataTypeTo(DspLong)
    TempDSPWave = DataWf.ConvertStreamTo(tldspParallel, TotallySize + 1, 0, Bit0IsMsb)
    TrimTemp = TrimCode + TempDSPWave.Element(0)
    
    If b_SetupToBit0 = False Or doallFlag = True Then
        For i = 0 To SizeTemp
            If TrimTemp <> 0 Then
                If TrimTemp \ (2 ^ (SizeTemp - i)) <> 0 Then
                    DataWf.Element(TotallySize - i) = 1
                Else
                    DataWf.Element(TotallySize - i) = 0
                End If
                If i <> SizeTemp Then
                    If TrimTemp >= (2 ^ (SizeTemp - i)) Then
                        TrimTemp = TrimTemp - (2 ^ (SizeTemp - i))
                    End If
                End If
            Else
                DataWf.Element(TotallySize - i) = 0
            End If
        Next i
    End If
End Function
