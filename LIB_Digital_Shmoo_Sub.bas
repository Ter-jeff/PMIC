Attribute VB_Name = "LIB_Digital_Shmoo_Sub"
Option Explicit

Public Function Search_Low2High_First_Pass(Shmoo_result_PF As String) As Integer
    Dim char_pt As String
    Dim max As Integer
    Dim point As Integer
    Dim i As Long
    On Error GoTo err1
    max = Len(Shmoo_result_PF)
    
    For i = 1 To max
        char_pt = Mid(Shmoo_result_PF, i, 1)
        If (char_pt = "P") Then point = i: i = max
        
    Next i
    Search_Low2High_First_Pass = point
    Exit Function
err1:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Search_High2Low_First_Pass(Shmoo_result_PF As String) As Integer
    Dim char_pt As String
    Dim max As Integer
    Dim point As Integer
    Dim i As Long
    
    On Error GoTo err1:
    max = Len(Shmoo_result_PF)

    For i = max To 1 Step -1
        char_pt = Mid(Shmoo_result_PF, i, 1)
        If (char_pt = "P") Then point = i: i = 1
        
    Next i
    
    Search_High2Low_First_Pass = point
    Exit Function
err1:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Search_HVCC(Shmoo_result_PF As String) As Integer
    Dim str_len, start_point As Integer
    Dim search_dif As Boolean
    Dim char_pt As String
    Dim report_point As Integer
    Dim i As Long
    On Error GoTo err1
    
    str_len = Len(Shmoo_result_PF)
    start_point = Search_Low2High_First_Pass(Shmoo_result_PF)
    search_dif = False
    report_point = 0
    For i = start_point To str_len
        char_pt = Mid(Shmoo_result_PF, i, 1)
        
        If Not (search_dif) Then
            If (char_pt = "F") Then
                report_point = i - 1
                search_dif = True
            End If
        Else
            If (char_pt = "P") Then
                report_point = -1
                i = str_len
            End If
        End If
    Next i
    If (report_point = 0) Then report_point = str_len: ReportHVCC = False
    Search_HVCC = report_point
    Exit Function
err1:
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function Search_LVCC(Shmoo_result_PF As String) As Integer
    Dim str_len, start_point As Integer
    Dim search_dif As Boolean
    Dim char_pt As String
    Dim report_point As Integer
    Dim i As Long
    
    On Error GoTo err1
    
    str_len = Len(Shmoo_result_PF)
    start_point = Search_High2Low_First_Pass(Shmoo_result_PF)
    search_dif = False
    report_point = 0
    
    For i = start_point To 1 Step -1
        char_pt = Mid(Shmoo_result_PF, i, 1)
        
        If Not (search_dif) Then
            If (char_pt = "F") Then
                report_point = i + 1
                search_dif = True
            End If
        Else
            If (char_pt = "P") Then
                report_point = -1
                i = 1
            End If
        End If
    Next i
    If (report_point = 0) Then report_point = 1: ReportLVCC = False
    Search_LVCC = report_point
    Exit Function
err1:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Search_VIH_LVCC(Shmoo_result_PF As String) As Integer
    'Report -1 is shmoo hole,report -2 is first point fail
    Dim str_len, start_point As Integer
    Dim search_dif As Boolean
    Dim char_pt As String
    Dim report_point As Integer
    Dim i As Long
    
    On Error GoTo err1
    str_len = Len(Shmoo_result_PF)
    start_point = str_len
    If (Mid(Shmoo_result_PF, start_point, 1) = "F") Then Search_VIH_LVCC = -2: Exit Function
    search_dif = False
    report_point = 0
    
    For i = start_point To 1 Step -1
        char_pt = Mid(Shmoo_result_PF, i, 1)
        
        If Not (search_dif) Then
            If (char_pt = "F") Then
                report_point = i + 1
                search_dif = True
            End If
        Else
            If (char_pt = "P") Then
                report_point = -1
                i = 1
            End If
        End If
    Next i
    If (report_point = 0) Then report_point = 1
    Search_VIH_LVCC = report_point
    Exit Function
err1:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Search_VIL_HVCC(Shmoo_result_PF As String) As Integer
    'Report -1 is shmoo hole,report -2 is first point fail
    Dim str_len, start_point As Integer
    Dim search_dif As Boolean
    Dim char_pt As String
    Dim report_point As Integer
    Dim i As Long
    
    On Error GoTo err1
    str_len = Len(Shmoo_result_PF)
    start_point = 1
    If (Mid(Shmoo_result_PF, 1, 1) = "F") Then Search_VIL_HVCC = -2: Exit Function
    search_dif = False
    report_point = 0
    For i = start_point To str_len
        char_pt = Mid(Shmoo_result_PF, i, 1)
        
        If Not (search_dif) Then
            If (char_pt = "F") Then
                report_point = i - 1
                search_dif = True
            End If
        Else
            If (char_pt = "P") Then
                report_point = -1
                i = str_len
            End If
        End If
    Next i
    If (report_point = 0) Then report_point = str_len
    Search_VIL_HVCC = report_point
    Exit Function
err1:
    If AbortTest Then Exit Function Else Resume Next
End Function
