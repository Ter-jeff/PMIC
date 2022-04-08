Attribute VB_Name = "LIB_Common_AP"
Option Explicit
Function IEDA_GetString(ByRef InputStr As String, RegistryName As String)
'(ByRef InputStr As String, FuseCategory As String, CategoryIndex As Integer)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "IEDA_GetString"
    Dim site As Variant
    Dim TmpString As String

        For Each site In TheExec.sites.Existing
                        Select Case RegistryName
                        'ECID IEDA
                           Case "eFuseLotNumber"
                                TmpString = ECIDFuse.Category(ECIDIndex("Lot_ID")).Read.ValStr(site)
                            Case "eFuseWaferID"
                                TmpString = ECIDFuse.Category(ECIDIndex("Wafer_ID")).Read.ValStr(site)
                            Case "eFuseDieX"
                                TmpString = ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.ValStr(site)
                            Case "eFuseDieY"
                                TmpString = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.ValStr(site)
                            Case "Hram_ECID_53bit"
                                TmpString = ECIDFuse.Category(gI_Index_DEID).Read.BitStrL(site)

''                            If CategoryIndex = gI_Index_53bits Then
''                                TmpString = ECIDFuse.Category(CategoryIndex).Read.BitStrL(Site)
''                            Else
''                                TmpString = ECIDFuse.Category(CategoryIndex).Read.ValStr(Site)
''                            End If
                          'UID IEDA
                        Case "Prov_Code"
                                Call IEDA_UID_Decode
                                If UIDFuse.Category(UIDIndex("UID_Code")).Read.ValStr(site) = "" Then
                                    TmpString = ""  'site not enable
                                ElseIf CDbl(UIDFuse.Category(UIDIndex("UID_Code")).Read.ValStr(site)) = 0 Then
                                    TmpString = "0"
                                ElseIf CDbl(UIDFuse.Category(UIDIndex("UID_Code")).LoLMT) < CDbl(UIDFuse.Category(UIDIndex("UID_Code")).Read.ValStr(site)) And CDbl(UIDFuse.Category(UIDIndex("UID_Code")).Read.ValStr(site)) < CDbl(UIDFuse.Category(UIDIndex("UID_Code")).HiLMT) Then
                                    TmpString = "1"
                                End If
                        'CFG IEDA
                        Case "SVM_CFuse"
                            TmpString = CFGFuse.Category(gI_CFG_firstbits_index).Read.BitStrM(site)
''                            If CategoryIndex = gI_CFG_firstbits_index Then
''                                TmpString = CFGFuse.Category(CategoryIndex).Read.BitStrM(Site)
''                            Else
''                                TmpString = CFGFuse.Category(CategoryIndex).Read.ValStr(Site)
''                            End If
                        Case "TMPS1_Untrim"
                            TmpString = gS_TMPS1_Untrim(site)
                        Case "TMPS2_Untrim"
                            TmpString = gS_TMPS2_Untrim(site)
                        Case "TMPS3_Untrim"
                            TmpString = gS_TMPS3_Untrim(site)
                        Case "TMPS4_Untrim"
                            TmpString = gS_TMPS4_Untrim(site)
                        Case "TMPS5_Untrim"
                            TmpString = gS_TMPS5_Untrim(site)
                        Case "TMPS6_Untrim"
                            TmpString = gS_TMPS6_Untrim(site)
                        Case "TMPS7_Untrim"
                            TmpString = gS_TMPS7_Untrim(site)
                        Case "TMPS8_Untrim"
                            TmpString = gS_TMPS8_Untrim(site)
                        Case "TMPS9_Untrim"
                            TmpString = gS_TMPS9_Untrim(site)
                        Case "TMPS10_Untrim"
                            TmpString = gS_TMPS10_Untrim(site)
                        Case "TMPS11_Untrim"
                            TmpString = gS_TMPS11_Untrim(site)
                        Case "TMPS12_Untrim"
                            TmpString = gS_TMPS12_Untrim(site)
                        Case "TMPS13_Untrim"
                            TmpString = gS_TMPS13_Untrim(site)
                        Case "TMPS14_Untrim"
                            TmpString = gS_TMPS14_Untrim(site)


                        Case "TMPS1_Trim"
                            TmpString = gS_TMPS1_Trim(site)
                        Case "TMPS2_Trim"
                            TmpString = gS_TMPS2_Trim(site)
                        Case "TMPS3_Trim"
                            TmpString = gS_TMPS3_Trim(site)
                        Case "TMPS4_Trim"
                            TmpString = gS_TMPS4_Trim(site)
                        Case "TMPS5_Trim"
                            TmpString = gS_TMPS5_Trim(site)
                        Case "TMPS6_Trim"
                            TmpString = gS_TMPS6_Trim(site)
                        Case "TMPS7_Trim"
                            TmpString = gS_TMPS7_Trim(site)
                        Case "TMPS8_Trim"
                            TmpString = gS_TMPS8_Trim(site)
                        Case "TMPS9_Trim"
                            TmpString = gS_TMPS9_Trim(site)
                        Case "TMPS10_Trim"
                            TmpString = gS_TMPS10_Trim(site)
                        Case "TMPS11_Trim"
                            TmpString = gS_TMPS11_Trim(site)
                        Case "TMPS12_Trim"
                            TmpString = gS_TMPS12_Trim(site)
                        Case "TMPS13_Trim"
                            TmpString = gS_TMPS13_Trim(site)
                        Case "TMPS14_Trim"
                            TmpString = gS_TMPS14_Trim(site)


                        Case "TMPS1"
                            TmpString = gS_TMPS1(site)
                        Case "TMPS2"
                            TmpString = gS_TMPS2(site)
                        Case "TMPS3"
                            TmpString = gS_TMPS3(site)
                        Case "TMPS4"
                            TmpString = gS_TMPS4(site)
                        Case "TMPS5"
                            TmpString = gS_TMPS5(site)
                        Case "TMPS6"
                            TmpString = gS_TMPS6(site)
                        Case "TMPS7"
                            TmpString = gS_TMPS7(site)
                        Case "TMPS8"
                            TmpString = gS_TMPS8(site)
                        Case "TMPS9"
                            TmpString = gS_TMPS9(site)
                        Case "TMPS10"
                            TmpString = gS_TMPS10(site)
                        Case "TMPS11"
                            TmpString = gS_TMPS11(site)
                        Case "TMPS12"
                            TmpString = gS_TMPS12(site)
                        Case "TMPS13"
                            TmpString = gS_TMPS13(site)
                        Case "TMPS14"
                            TmpString = gS_TMPS14(site)
                                                Case "BKM"
                            TmpString = CFGFuse.Category(CFGIndex("bkm_package")).Write.Decimal(site)
                        Case "BKM_Fuse"
                            TmpString = CFGFuse.Category(CFGIndex("bkm_package")).Read.Decimal(site)

''                        Case "UDR"
''                            TmpString = UDRFuse.Category(CategoryIndex).Read.ValStr(Site)
''                        Case "SEN"
''                            TmpString = SENFuse.Category(CategoryIndex).Read.ValStr(Site)

                        Case Else
                            TheExec.Datalog.WriteComment "print: warnining, no suitable registry choosed in VBT 'IEDA_GetString'."
                        End Select
            If (site = TheExec.sites.Existing.Count - 1) Then
                InputStr = InputStr + TmpString
            Else
                 InputStr = InputStr + TmpString + ","
            End If

        Next site

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Function IEDA_AutoCheck_Print(ByRef InputStr As String, RegistryName As String, DebugPrint As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "IEDA_AutoCheck_Print"
    Dim TmpString As String

    InputStr = auto_checkIEDAString(InputStr)
    If DebugPrint Then TheExec.Datalog.WriteComment "print: Set IEDA registry ( " & RegistryName & " ) = " & InputStr

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Public Function IEDA_UID_Decode(Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList)
'''
'''On Error GoTo errHandler
'''    Dim funcName As String:: funcName = "IEDA_UID_Decode"
'''
'''    Dim site As Variant
'''    Dim SingleBitArray() As Long, SingleBitSum As Long
'''    Dim DoubleBitArray() As Long, DoubleBitSum As Long
'''    Dim m_siteVar As String
'''    m_siteVar = "UIDChk_Var"
'''
'''    For Each site In TheExec.sites
'''        '''''initialize per Site
'''        ReDim SingleBitArray(UIDTotalBits - 1)
'''        ReDim DoubleBitArray(UIDBitsPerBlockUsed - 1)
'''        SingleBitSum = 0
'''        DoubleBitSum = 0
'''
'''            Call auto_OR_2Blocks("UID", gS_SingleStrArray, SingleBitArray, DoubleBitArray)  ''''to get gL_UID_FBC()
'''
'''            If (DisplayUID = True) Then
'''                TheExec.Datalog.WriteComment ""
'''                TheExec.Datalog.WriteComment "Read AES/UID data from DSSC at Site (" + CStr(site) + ")"
'''                Call auto_PrintAllBitbyDSSC(SingleBitArray, UIDReadCycle, UIDTotalBits, UIDReadBitWidth)
'''            End If
'''
'''            Call auto_Decode_UIDBinary_Data(DoubleBitArray)  ''''20150616 New
'''
'''    Next site
'''
'''    'TheExec.Flow.TestLimit resultVal:=gL_AES_FBC, lowVal:=0, hiVal:=0, Tname:="FailBitCount"
'''
'''Exit Function
'''
'''errHandler:
'''    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'''    If AbortTest Then Exit Function Else Resume Next
'''
End Function


