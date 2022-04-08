Attribute VB_Name = "LIB_Common_Pattern_Freq"
Option Explicit
'Revision History:
'V0.0 initial bring up
Public Function PatExculdePath(Pat As Variant) As String
    Dim patt_ary_temp() As String
    patt_ary_temp = Split(Pat, "\")
    PatExculdePath = patt_ary_temp(UBound(patt_ary_temp))
End Function

'*****************************************
'******        pattern set, patterns******
'*****************************************
' Decompose patset recursively and return a string "patt" with a list of .pat
Public Function PatSetToPat(ByVal patset As Pattern) As String
'   pat1    pat1a,pat1b,pat1c.pat
'   pat1a   pat1a.pat
'   pat1b   pat1b1, pat1b2
'   pat1b1  pat1b1.pat
'   pat1b2  pat1b2.pat

    Dim Pat_ary() As String, PatCnt As Long
    Dim Pat_ary1() As String, PatCnt1 As Long
    Dim patset_ary() As String, i As Long
    Dim patset1 As New Pattern
    Dim patt_str As String
    Dim patt As String
    
    patset_ary = Split(patset.Value, ",")
    patt = ""
    For i = 0 To UBound(patset_ary)
        Current_Patterns = ""
        Call PatsetDecompose(patset_ary(i))
        If patt <> "" Then
            patt = patt & "," & Current_Patterns
        Else
            patt = Current_Patterns
        End If
    Next i
    PatSetToPat = patt
End Function

Public Function PatSetToPat_EFuse(ByVal patset As Pattern, ByRef patt As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "PatSetToPat"

    Dim Pat_ary() As String, PatCnt As Long

    Pat_ary = TheExec.DataManager.Raw.GetPatternsInSet(patset, PatCnt)
    patset.Value = Pat_ary(0)
    While Not (LCase(Pat_ary(0)) Like "*.pat*")
        Pat_ary = TheExec.DataManager.Raw.GetPatternsInSet(patset, PatCnt)
        If PatCnt > 1 Then TheExec.ErrorLogMessage (patset & " is with more than one pattern in the pattern set")
        patset.Value = Pat_ary(0)
    Wend
    patt = Pat_ary(0)
    TheHdw.Patterns(patt).Load

Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetPatListFromPatternSet(TestPat As String, _
                              rtnPatNames() As String, _
                              rtnPatCnt As Long) As Boolean
    'used to be GetPatFromPatternSet
    Dim patt_list As String
    Dim patt As New Pattern
    On Error GoTo errHandler
    
    patt.Value = TestPat
    patt_list = PatSetToPat(patt)
    rtnPatNames = Split(patt_list, ",")
    rtnPatCnt = UBound(rtnPatNames) + 1
    If (UBound(rtnPatNames) >= 0) Then
        If LCase(rtnPatNames(0)) Like "*.pat*" Then
            GetPatListFromPatternSet = True
        End If
    End If

    Exit Function
    
Exit Function
errHandler:
    GetPatListFromPatternSet = False
    rtnPatCnt = -1

                If AbortTest Then Exit Function Else Resume Next
End Function
' do not use: only as the sub function recursively called in the PatSetToPat()
Public Function PatsetDecompose(PatSetName As String) As String
    Dim PatCnt As Long                          '<- Number of patterns in set
    Dim RawNameData() As String                 '<- Raw pattern name data
    Dim pIndex As Long
    Dim patt_str As String
    
    RawNameData = TheExec.DataManager.Raw.GetPatternsInSet(PatSetName, PatCnt)
    PatCnt = UBound(RawNameData)
    For pIndex = 0 To PatCnt
        If InStr(1, RawNameData(pIndex), ".pat", vbTextCompare) Then
            If Current_Patterns = "" Then
                Current_Patterns = RawNameData(pIndex)
            Else
                Current_Patterns = Current_Patterns & "," & RawNameData(pIndex)
            End If
        Else
            Call PatsetDecompose(RawNameData(pIndex))
        End If
    Next pIndex
End Function

'*****************************************
'******         frequecy measurement******
'*****************************************
Public Function Freq_MeasFreqSetup(Pin As PinList, Interval As Double, Optional MeasF_EventSource As FreqCtrEventSrcSel = 1)
    With TheHdw.Digital.Pins(Pin).FreqCtr
        .EventSource = MeasF_EventSource '' VOH
        .EventSlope = Positive
        .Interval = Interval
        .Enable = IntervalEnable
        .Clear
    End With
End Function

Public Function Freq_MeasFreqStart(Pin As PinList, Interval As Double, freq As PinListData)
    Dim CounterValue As New PinListData
    Dim site As Variant
    TheHdw.Digital.Pins(Pin).FreqCtr.Clear
    TheHdw.Digital.Pins(Pin).FreqCtr.start
    
    For Each site In TheExec.sites
        CounterValue = TheHdw.Digital.Pins(Pin).FreqCtr.Read
        freq = CounterValue.Math.Divide(Interval)
    Next site
End Function
