Attribute VB_Name = "LIB_EFUSE_Common"

Option Explicit

Private gDictIndex As New Dictionary
Private gDictIndex_ECID As New Dictionary
Private gDictIndex_CFG As New Dictionary
Private gDictIndex_CFGTab As New Dictionary
Private gDictIndex_CFGCond As New Dictionary
Private gDictIndex_UID As New Dictionary
Private gDictIndex_SEN As New Dictionary
Private gDictIndex_MON As New Dictionary
Private gDictIndex_UDR As New Dictionary
Private gDictIndex_UDRE As New Dictionary
Private gDictIndex_UDRP As New Dictionary
Private gDictIndex_CMP As New Dictionary
Private gDictIndex_CMPE As New Dictionary
Private gDictIndex_CMPP As New Dictionary

Public gB_newDlog_Flag As Boolean

'''20191230, move to VBT_LIB_MBIST
'Public DebugPrtImm As Boolean
'Public DebugPrtDlog As Boolean

Public gB_findECID_flag As Boolean
Public gB_findCFG_flag As Boolean
Public gB_findUID_flag As Boolean
Public gB_findUDR_flag As Boolean
Public gB_findSEN_flag As Boolean
Public gB_findMON_flag As Boolean
Public gB_findCMP_flag As Boolean
Public gB_findCFGTable_flag As Boolean
Public gB_findCFGCondTable_flag As Boolean ''''for new CFG_Condition_Table, 20170630
Public gB_findUDRE_flag As Boolean
Public gB_findUDRP_flag As Boolean
Public gB_findCMPE_flag As Boolean
Public gB_findCMPP_flag As Boolean

Public ECIDFuse As EFuseCategorySyntax
Public CFGFuse As EFuseCategorySyntax
Public UIDFuse As EFuseCategorySyntax
Public UDRFuse As EFuseCategorySyntax
Public SENFuse As EFuseCategorySyntax
Public MONFuse As EFuseCategorySyntax
Public CMPFuse As EFuseCategorySyntax
Public UDRE_Fuse As EFuseCategorySyntax
Public UDRP_Fuse As EFuseCategorySyntax
Public CMPE_Fuse As EFuseCategorySyntax
Public CMPP_Fuse As EFuseCategorySyntax

Public CFGTable As ConfigTableSyntax ''''20170630 update
''Public CFGCond As CFGCondTableSyntax

Public CRC_Shift_Out_String As New SiteVariant  ''''20170623 added
Public gL_CFG_Cond_beforeWrite_BitVal() As Long ''''20170717 add, used for the simulation before Write
Public gL_CFG_Cond_allBitWidth As Long          ''''20170911 add
Public gL_CFG_Cond_min_lsbbit As Long           ''''20170911 add
Public gL_CFG_Cond_max_msbbit As Long           ''''20170911 add
Public gL_CFG_Cond_BitIndex() As Long           ''''20180917 add



''''201811XX
' Function to retrieve a Index. This
' function can be called by user interpose functions to access
' previously stored measurements
Public Function eFuse_GetStoredIndex(FuseType As eFuseBlockType, ByVal KeyName As String) As Variant
On Error GoTo errHandler
    Dim funcName As String:: funcName = "eFuse_GetStoredIndex"

    Select Case FuseType
        Case eFuse_ECID:
            Set gDictIndex = gDictIndex_ECID
        Case eFuse_CFG:
            Set gDictIndex = gDictIndex_CFG
        Case eFuse_CFGTab:
            Set gDictIndex = gDictIndex_CFGTab
        Case eFuse_CFGCond:
            Set gDictIndex = gDictIndex_CFGCond
        Case eFuse_UDR:
            Set gDictIndex = gDictIndex_UDR
        Case eFuse_UDRE:
            Set gDictIndex = gDictIndex_UDRE
        Case eFuse_UDRP:
            Set gDictIndex = gDictIndex_UDRP
        Case eFuse_MON:
            Set gDictIndex = gDictIndex_MON
    End Select

    KeyName = LCase(KeyName)
    If Not gDictIndex.Exists(KeyName) Then
        eFuse_GetStoredIndex = -1
         If Not (KeyName Like "ids_*_85") Then
        TheExec.ErrorLogMessage "Stored Index of " & KeyName & " not found"
        End If
    Else
        eFuse_GetStoredIndex = gDictIndex(KeyName)
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''201811XX
' Function to store a Index for later retrieval,
' typically from a custom user interpose function
Public Function eFuse_AddStoredIndex(FuseType As eFuseBlockType, ByVal KeyName As String, ByRef obj As Variant)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "eFuse_AddStoredIndex"

    Select Case FuseType
        Case eFuse_ECID:
            Set gDictIndex = gDictIndex_ECID
        Case eFuse_CFG:
            Set gDictIndex = gDictIndex_CFG
        Case eFuse_CFGTab:
            Set gDictIndex = gDictIndex_CFGTab
        Case eFuse_CFGCond:
            Set gDictIndex = gDictIndex_CFGCond
        Case eFuse_UDR:
            Set gDictIndex = gDictIndex_UDR
        Case eFuse_UDRP:
            Set gDictIndex = gDictIndex_UDRP
        Case eFuse_UDRE:
            Set gDictIndex = gDictIndex_UDRE
        Case eFuse_MON:
            Set gDictIndex = gDictIndex_MON
    End Select
    
    KeyName = LCase(KeyName)
    If gDictIndex.Exists(KeyName) Then
        gDictIndex.Remove (KeyName)
    End If
    gDictIndex.Add KeyName, obj
    TheHdw.Wait 1 * us
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Protect_eFuse_Sheet()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "Protect_eFuse_Sheet"

    Sheets(gS_eFuse_sheetName).Protect
    
    ''''20170220 update
'    If (UCase(gS_EFuse_Orientation) = UCase("SingleUp")) Then
'        Sheets("eFuse_ECID_Bit_Allocation_1Bit").Protect
'    ElseIf (UCase(gS_EFuse_Orientation) = "RIGHT2LEFT") Then
'        Sheets("eFuse_ECID_Bit_Allocation_2Bit").Protect
'    End If

    ''''20170220 update
    If (gB_findCFG_flag) Then
        Sheets(gS_cfgTable_sheetName).Protect
        Sheets(gS_cfgTable_SVM_sheetName).Protect
    End If

    '========================================
    '=  How to unprotect the spread sheet   =
    '========================================
    'Step1.
        'Select the spread sheet of "Config_Table"
    'Step2.
        'Switch the VBT window and launch the Immediate Window
    'Step3.
        'In the Immediate Window, key in the following command
        '===>  activesheet.unprotect

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function UnProtect_eFuse_Sheet()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "UnProtect_eFuse_Sheet"

    Sheets(gS_eFuse_sheetName).Unprotect
    ''Sheets(gS_cfgTable_sheetName).Unprotect
    ''Sheets(gS_cfgTable_SVM_sheetName).Unprotect

    '========================================
    '=  How to unprotect the spread sheet   =
    '========================================
    'Step1.
        'Select the spread sheet of "Config_Table"
    'Step2.
        'Switch the VBT window and launch the Immediate Window
    'Step3.
        'In the Immediate Window, key in the following command
        '===>  activesheet.unprotect

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function PrintDataLog(myStr As String)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "PrintDataLog"
    
    TheExec.Datalog.WriteComment myStr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function DebugPrintLog(myStr As String)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "DebugPrintLog"
    
    ''''Print in the Immediate Window
    If (DebugPrtImm) Then
        Debug.Print myStr
    End If
    
    ''''Print in the Datalog
    If (DebugPrtDlog) Then
        TheExec.Datalog.WriteComment myStr
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


''''----------------------------------------------------------------
''''20151008 update for New Format of 'EFUSE_BitDef_Table' sheet
''''20160831 update to add CMPFuse
''''20171103 update syntax with bank_* eFuse Dit Def
''''         add UDR_E and UDR_P
''''----------------------------------------------------------------
Public Function parse_eFuse_ChkList_New(sheetName As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "parse_eFuse_ChkList_New"
    
    Dim mysheet As Worksheet
    Dim myCell As Object
    Dim offCell As Object
    Dim i As Long, j As Long, k As Long
    Dim m_cellStr As String
    Const m_CateNumSize = 4096 ''''20170925 add, 201811XX
    Dim m_keyname As String
    
    DebugPrtImm = False
    DebugPrtDlog = False

    Set mysheet = Sheets(sheetName)
    
    Set myCell = mysheet.range("A1")
    m_cellStr = UCase(Trim(myCell.Value))
    
    ''''Get the parameters for the "Single/Double-Bit Orientation"
    Do While m_cellStr <> "END"
        DebugPrintLog "0...input myCell=" & myCell

        ''''New EFuse ChkList Table Format
        If (m_cellStr Like UCase("*Right-to-Left*") Or m_cellStr Like UCase("*2*Bit*")) Then
            gS_EFuse_Orientation = "RIGHT2LEFT"
            gE_eFuse_Orientation = eFuse_2_Bit ''''20180628 add
            DebugPrintLog "     Single-Bit Double-Bit Orientation = " + gS_EFuse_Orientation
            Exit Do

        ElseIf (m_cellStr Like UCase("*Up-to-Down*")) Then
            gS_EFuse_Orientation = "UP2DOWN"
            gE_eFuse_Orientation = eFuse_UP2DOWN ''''20180628 add
            DebugPrintLog "     Single-Bit Double-Bit Orientation = " + gS_EFuse_Orientation
            Exit Do

        ElseIf (m_cellStr Like UCase("*Single-Up*") Or m_cellStr Like UCase("*1*Bit*")) Then
            gS_EFuse_Orientation = "SingleUp"
            gE_eFuse_Orientation = eFuse_1_Bit ''''20180628 add
            DebugPrintLog "     Single-Bit Orientation = " + gS_EFuse_Orientation
            Exit Do

        ElseIf (m_cellStr Like UCase("*Single-Down*")) Then
            gS_EFuse_Orientation = "SingleDown"
            gE_eFuse_Orientation = eFuse_SingleDown ''''20180628 add
            DebugPrintLog "     Single-Bit Orientation = " + gS_EFuse_Orientation
            Exit Do

        ElseIf (m_cellStr Like UCase("*Single-Right*")) Then
            gS_EFuse_Orientation = "SingleRight"
            gE_eFuse_Orientation = eFuse_SingleRight ''''20180628 add
            DebugPrintLog "     Single-Bit Orientation = " + gS_EFuse_Orientation
            Exit Do

        ElseIf (m_cellStr Like UCase("*Single-Left*")) Then
            gS_EFuse_Orientation = "SingleLeft"
            gE_eFuse_Orientation = eFuse_SingleLeft ''''20180628 add
            DebugPrintLog "     Single-Bit Orientation = " + gS_EFuse_Orientation
            Exit Do

        Else
            gS_EFuse_Orientation = "Unknown"
            gE_eFuse_Orientation = eFuse_Orient_Unknown ''''20180628 add
            DebugPrintLog "     Orientation = " + gS_EFuse_Orientation
            Exit Do

        End If

        ''''if cell search from up   to down,  (rowOffset:=1, columnOffset:=0)
        ''''if cell search from left to right, (rowOffset:=0, columnOffset:=1)
        Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
        m_cellStr = UCase(Trim(myCell.Value))
    Loop
    gDL_eFuse_Orientation = gE_eFuse_Orientation
    gE_eFuse_Orientation_1Bits = eFuse_1_Bit ''20190801
    
    ''''---------------------------------------------------------------------------------------------------------
    ''''Get the parameters for the "*eFuse Bit Def*"
    Dim offcolStr As String
    Dim M As Integer
    Dim idx_seqSTART As Integer
    Dim idx_seqEnd As Integer
    Dim idx_MSBbit As Integer
    Dim idx_LSBbit As Integer
    Dim idx_BitWidth As Integer
    Dim naCnt As Integer
    Dim idx_NA1 As Integer
    Dim idx_NA2 As Integer
    Dim idx_NA3 As Integer
    Dim idx_NA4 As Integer
    Dim idx_USI_MSBbitCycle As Integer
    Dim idx_USI_LSBbitCycle As Integer
    Dim idx_USO_MSBbitCycle As Integer
    Dim idx_USO_LSBbitCycle As Integer
    Dim idx_Stage As Integer
    Dim idx_LoLMT As Integer
    Dim idx_HiLMT As Integer
    Dim idx_Resolution As Integer
    Dim idx_Algorithm As Integer
    Dim idx_Comment As Integer
    Dim idx_DefaultOrReal As Integer
    Dim idx_DefaultValue As Integer
    Dim idx_Difference As Integer

    Dim n As Integer
    Dim findAllParamHeader As Boolean
    Dim findECIDHeader As Boolean
    Dim findCFGHeader As Boolean
    Dim findUIDHeader As Boolean
    Dim findUDRHeader As Boolean
    Dim findSENHeader As Boolean
    Dim findMONHeader As Boolean
    Dim findCMPHeader As Boolean
    Dim findUDRE_Header As Boolean
    Dim findUDRP_Header As Boolean
    Dim findCMPE_Header As Boolean
    Dim findCMPP_Header As Boolean
    
    Dim CateCnt_ECID As Integer
    Dim CateCnt_CFG As Integer
    Dim CateCnt_UID As Integer
    Dim CateCnt_UDR As Integer
    Dim CateCnt_SEN As Integer
    Dim CateCnt_MON As Integer
    Dim CateCnt_CMP As Integer
    Dim CateCnt_UDRE As Integer
    Dim CateCnt_UDRP As Integer
    Dim CateCnt_CMPE As Integer
    Dim CateCnt_CMPP As Integer

    findAllParamHeader = False
    findECIDHeader = False
    findCFGHeader = False
    findUIDHeader = False
    findUDRHeader = False
    findSENHeader = False
    findMONHeader = False
    findCMPHeader = False
    findUDRE_Header = False
    findUDRP_Header = False
    findCMPE_Header = False
    findCMPP_Header = False

    gB_findECID_flag = False
    gB_findCFG_flag = False
    gB_findUID_flag = False
    gB_findUDR_flag = False
    gB_findSEN_flag = False
    gB_findMON_flag = False
    gB_findCMP_flag = False
    gB_findUDRE_flag = False
    gB_findUDRP_flag = False
    gB_findCMPE_flag = False
    gB_findCMPP_flag = False

    CateCnt_ECID = 0
    CateCnt_CFG = 0
    CateCnt_UDR = 0
    CateCnt_UID = 0
    CateCnt_SEN = 0
    CateCnt_MON = 0
    CateCnt_CMP = 0
    CateCnt_UDRE = 0
    CateCnt_UDRP = 0
    CateCnt_CMPE = 0
    CateCnt_CMPP = 0

    Set myCell = mysheet.range("A1")
    m_cellStr = UCase(Trim(myCell.Value))
    DebugPrintLog "0...input myCell=" + CStr(myCell.Value)
    
    Do While (m_cellStr <> "END")
      
        If ((m_cellStr Like UCase("Direct*Access*Mode*")) Or (m_cellStr Like UCase("JTAG*Access*Mode*")) Or (m_cellStr Like UCase("Read*Out*from*JTAG*OUT*")) Or (m_cellStr Like UCase("Read*Out*from*JTAG*TDO*"))) Then

            If (findECIDHeader) Then
                DebugPrintLog "1...ECIDFuse.Category Array Size = " + CStr(CateCnt_ECID)
                If (CateCnt_ECID > 0) Then
                    gB_findECID_flag = True
                    ReDim Preserve ECIDFuse.Category(CateCnt_ECID - 1) ''''final dimension
                End If
                '''--------------------------------------------------------------------------
                findECIDHeader = False ''''MUST have to know that ECID process is done
            End If

            If (findCFGHeader) Then
                DebugPrintLog "2...CFGFuse.Category Array Size = " + CStr(CateCnt_CFG)
                If (CateCnt_CFG > 0) Then
                    gB_findCFG_flag = True
                    ReDim Preserve CFGFuse.Category(CateCnt_CFG - 1) ''''final dimension
                End If
                findCFGHeader = False ''''MUST have to know that CFG process is done
            End If

            If (findUIDHeader) Then
                DebugPrintLog "3......UIDFuse.Category Array Size = " + CStr(CateCnt_UID)
                If (CateCnt_UID > 0) Then
                    gB_findUID_flag = True
                    ReDim Preserve UIDFuse.Category(CateCnt_UID - 1) ''''final dimension
                End If
                findUIDHeader = False ''''MUST have to know that UID process is done
            End If
           
            If (findUDRHeader) Then
                DebugPrintLog "4...UDRFuse.Category Array Size = " + CStr(CateCnt_UDR)
                If (CateCnt_UDR > 0) Then
                    gB_findUDR_flag = True
                    ReDim Preserve UDRFuse.Category(CateCnt_UDR - 1) ''''final dimension
                End If
                findUDRHeader = False ''''MUST have to know that UDR process is done
            End If

            If (findSENHeader) Then
                DebugPrintLog "5...SENFuse.Category Array Size = " + CStr(CateCnt_SEN)
                If (CateCnt_SEN > 0) Then
                    gB_findSEN_flag = True
                    ReDim Preserve SENFuse.Category(CateCnt_SEN - 1) ''''final dimension
                End If
                findSENHeader = False ''''MUST have to know that SEN process is done
            End If
            
            If (findMONHeader) Then
                DebugPrintLog "5...MONFuse.Category Array Size = " + CStr(CateCnt_MON)
                If (CateCnt_MON > 0) Then
                    gB_findMON_flag = True
                    ReDim Preserve MONFuse.Category(CateCnt_MON - 1) ''''final dimension
                End If
                findMONHeader = False ''''MUST have to know that MON process is done
            End If

            If (findCMPHeader) Then
                DebugPrintLog "5...CMPFuse.Category Array Size = " + CStr(CateCnt_CMP)
                If (CateCnt_CMP > 0) Then
                    gB_findCMP_flag = True
                    ReDim Preserve CMPFuse.Category(CateCnt_CMP - 1) ''''final dimension
                End If
                findCMPHeader = False ''''MUST have to know that CMP process is done
            End If

            ''''20171103 add
            If (findUDRE_Header) Then
                DebugPrintLog "6...UDRE_Fuse.Category Array Size = " + CStr(CateCnt_UDRE)
                If (CateCnt_UDRE > 0) Then
                    gB_findUDRE_flag = True
                    ReDim Preserve UDRE_Fuse.Category(CateCnt_UDRE - 1) ''''final dimension
                End If
                findUDRE_Header = False ''''MUST have to know that UDRE process is done
            End If

            If (findUDRP_Header) Then
                DebugPrintLog "7...UDRP_Fuse.Category Array Size = " + CStr(CateCnt_UDRP)
                If (CateCnt_UDRP > 0) Then
                    gB_findUDRP_flag = True
                    ReDim Preserve UDRP_Fuse.Category(CateCnt_UDRP - 1) ''''final dimension
                End If
                findUDRP_Header = False ''''MUST have to know that UDRP process is done
            End If

            If (findCMPE_Header) Then
                DebugPrintLog "8...CMPE_Fuse.Category Array Size = " + CStr(CateCnt_CMPE)
                If (CateCnt_CMPE > 0) Then
                    gB_findCMPE_flag = True
                    ReDim Preserve CMPE_Fuse.Category(CateCnt_CMPE - 1) ''''final dimension
                End If
                findCMPE_Header = False ''''MUST have to know that CMPE process is done
            End If

            If (findCMPP_Header) Then
                DebugPrintLog "9...CMPP_Fuse.Category Array Size = " + CStr(CateCnt_CMPP)
                If (CateCnt_CMPP > 0) Then
                    gB_findCMPP_flag = True
                    ReDim Preserve CMPP_Fuse.Category(CateCnt_CMPP - 1) ''''final dimension
                End If
                findCMPP_Header = False ''''MUST have to know that CMPP process is done
            End If

            ''''------------
            GoTo nextRow
            ''''------------
        End If
        
        If (m_cellStr Like UCase("*ECID*Bit*Def*")) Then
            n = 0
            findECIDHeader = True
            findCFGHeader = False
            findUIDHeader = False
            findUDRHeader = False
            findSENHeader = False
            findMONHeader = False
            findCMPHeader = False
            findUDRE_Header = False
            findUDRP_Header = False
            findCMPE_Header = False
            findCMPP_Header = False
            ReDim ECIDFuse.Category(m_CateNumSize) ''''initialize

            ''''Table Column Sequence::
            ''''ECID eFuse Bit Def, Start Bit, End Bit, Bit Width, N/A, N/A, N/A, N/A, programming stage, Low Limit, High Limit, IDS Resolution, Algorithm, Description, Default or Real, Default Value, Difference
            ''''UDR eFuse Bit Def,  LSB Bit,   MSB Bit, Bit Width, USI LSB-Bit Cycle, USI MSB-Bit Cycle, USO LSB-Bit Cycle, USO MSB-Bit Cycle, programming stage, Low Limit, High Limit, IDS Resolution, Algorithm, Description, Default or Real, Default Value, Difference
            
            ''''-----------------------------------------------------------------------------------------------------------
            ''''Get the Specific Index
            M = 0: naCnt = 0
            offcolStr = ""
            findAllParamHeader = False

            ''''Initialize
            idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
            idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
            idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
            idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
            idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
            idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1

            Do
                M = M + 1
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                offcolStr = UCase(Trim(offCell.Value))
                DebugPrintLog "1...ECID input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                
                If (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Start*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                    idx_seqSTART = M
                    idx_MSBbit = M
                ElseIf (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("End*Bit") Or offcolStr Like UCase("Seq*End")) Then
                    idx_seqEnd = M
                    idx_LSBbit = M
                ElseIf (offcolStr Like UCase("Bit*Width")) Then
                    idx_BitWidth = M
                ElseIf (offcolStr = UCase("N/A")) Then
                    naCnt = naCnt + 1
                    Select Case (naCnt)
                    Case 1
                        idx_NA1 = M
                    Case 2
                        idx_NA2 = M
                    Case 3
                        idx_NA3 = M
                    Case 4
                        idx_NA4 = M
                    Case Else
                    End Select
                ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                    idx_USI_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                    idx_USI_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                    idx_USO_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                    idx_USO_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("program*stage")) Then
                    idx_Stage = M
                ElseIf (offcolStr Like UCase("Low*Limit")) Then
                    idx_LoLMT = M
                ElseIf (offcolStr Like UCase("High*Limit")) Then
                    idx_HiLMT = M
                ElseIf (offcolStr Like UCase("*Resolution")) Then
                    idx_Resolution = M
                ElseIf (offcolStr Like UCase("Algorithm")) Then
                    idx_Algorithm = M
                ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                    idx_Comment = M
                ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                    idx_DefaultOrReal = M
                ElseIf (offcolStr Like UCase("Default*Value")) Then
                    idx_DefaultValue = M
                ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                    idx_Difference = M
                    findAllParamHeader = True
                    Exit Do
                End If
            Loop While (findAllParamHeader = False)
            ''''-----------------------------------------------------------------------------------------------------------
            GoTo nextRow

        ElseIf (m_cellStr Like UCase("*Config*Bit*Def*")) Then

            ReDim CFGFuse.Category(m_CateNumSize) ''''initialize
            n = 0
            findECIDHeader = False
            findCFGHeader = True
            findUIDHeader = False
            findUDRHeader = False
            findSENHeader = False
            findMONHeader = False
            findCMPHeader = False
            findUDRE_Header = False
            findUDRP_Header = False
            findCMPE_Header = False
            findCMPP_Header = False

            ''''-----------------------------------------------------------------------------------------------------------
            ''''Get the Specific Index
            M = 0: naCnt = 0
            offcolStr = ""
            findAllParamHeader = False

            ''''Initialize
            idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
            idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
            idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
            idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
            idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
            idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1

            Do
                M = M + 1
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                offcolStr = UCase(Trim(offCell.Value))
                DebugPrintLog "2... CFG input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                
                If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                    idx_LSBbit = M
                ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                    idx_MSBbit = M
                ElseIf (offcolStr Like UCase("Bit*Width")) Then
                    idx_BitWidth = M
                ElseIf (offcolStr = UCase("N/A")) Then
                    naCnt = naCnt + 1
                    Select Case (naCnt)
                    Case 1
                        idx_NA1 = M
                    Case 2
                        idx_NA2 = M
                    Case 3
                        idx_NA3 = M
                    Case 4
                        idx_NA4 = M
                    Case Else
                    End Select
                ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                    idx_USI_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                    idx_USI_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                    idx_USO_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                    idx_USO_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("program*stage")) Then
                    idx_Stage = M
                ElseIf (offcolStr Like UCase("Low*Limit")) Then
                    idx_LoLMT = M
                ElseIf (offcolStr Like UCase("High*Limit")) Then
                    idx_HiLMT = M
                ElseIf (offcolStr Like UCase("*Resolution")) Then
                    idx_Resolution = M
                ElseIf (offcolStr Like UCase("Algorithm")) Then
                    idx_Algorithm = M
                ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                    idx_Comment = M
                ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                    idx_DefaultOrReal = M
                ElseIf (offcolStr Like UCase("Default*Value")) Then
                    idx_DefaultValue = M
                ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                    idx_Difference = M
                    findAllParamHeader = True
                    Exit Do
                End If
            Loop While (findAllParamHeader = False)
            ''''-----------------------------------------------------------------------------------------------------------
            GoTo nextRow
            
         ElseIf (m_cellStr Like UCase("*UID*Bit*Def*")) Then

            ReDim UIDFuse.Category(m_CateNumSize) ''''initialize
            n = 0
            findECIDHeader = False
            findCFGHeader = False
            findUIDHeader = True
            findUDRHeader = False
            findSENHeader = False
            findMONHeader = False
            findCMPHeader = False
            findUDRE_Header = False
            findUDRP_Header = False
            findCMPE_Header = False
            findCMPP_Header = False
            
            ''''-----------------------------------------------------------------------------------------------------------
            ''''Get the Specific Index
            M = 0: naCnt = 0
            offcolStr = ""
            findAllParamHeader = False

            ''''Initialize
            idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
            idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
            idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
            idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
            idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
            idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1

            Do
                M = M + 1
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                offcolStr = UCase(Trim(offCell.Value))
                DebugPrintLog "3... UID input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                
                If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                    idx_LSBbit = M
                ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                    idx_MSBbit = M
                ElseIf (offcolStr Like UCase("Bit*Width")) Then
                    idx_BitWidth = M
                ElseIf (offcolStr = UCase("N/A")) Then
                    naCnt = naCnt + 1
                    Select Case (naCnt)
                    Case 1
                        idx_NA1 = M
                    Case 2
                        idx_NA2 = M
                    Case 3
                        idx_NA3 = M
                    Case 4
                        idx_NA4 = M
                    Case Else
                    End Select
                ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                    idx_USI_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                    idx_USI_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                    idx_USO_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                    idx_USO_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("program*stage")) Then
                    idx_Stage = M
                ElseIf (offcolStr Like UCase("Low*Limit")) Then
                    idx_LoLMT = M
                ElseIf (offcolStr Like UCase("High*Limit")) Then
                    idx_HiLMT = M
                ElseIf (offcolStr Like UCase("*Resolution")) Then
                    idx_Resolution = M
                ElseIf (offcolStr Like UCase("Algorithm")) Then
                    idx_Algorithm = M
                ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                    idx_Comment = M
                ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                    idx_DefaultOrReal = M
                ElseIf (offcolStr Like UCase("Default*Value")) Then
                    idx_DefaultValue = M
                ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                    idx_Difference = M
                    findAllParamHeader = True
                    Exit Do
                End If
            Loop While (findAllParamHeader = False)
            ''''-----------------------------------------------------------------------------------------------------------
            GoTo nextRow

        ''''20171103 add
        ElseIf (m_cellStr Like UCase("*UDR_E*Bit*Def*")) Then

            ''''20180104 update per customer format request
            If (m_cellStr Like UCase("*CMP*Bit*Def*")) Then ''''case "bank_UDR_E CMP eFuse Bit Def"
                ReDim CMPE_Fuse.Category(m_CateNumSize) ''''initialize
                n = 0
                findECIDHeader = False
                findCFGHeader = False
                findUIDHeader = False
                findUDRHeader = False
                findSENHeader = False
                findMONHeader = False
                findCMPHeader = False
                findUDRE_Header = False
                findUDRP_Header = False
                findCMPE_Header = True
                findCMPP_Header = False
                ''''-----------------------------------------------------------------------------------------------------------
                ''''Get the Specific Index
                M = 0: naCnt = 0
                offcolStr = ""
                findAllParamHeader = False
    
                ''''Initialize
                idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
                idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
                idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
                idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
                idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
                idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1
    
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    offcolStr = UCase(Trim(offCell.Value))
                    DebugPrintLog "5... CMPE input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                    
                    If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                        idx_LSBbit = M
                    ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                        idx_MSBbit = M
                    ElseIf (offcolStr Like UCase("Bit*Width")) Then
                        idx_BitWidth = M
                    ElseIf (offcolStr = UCase("N/A")) Then
                        naCnt = naCnt + 1
                        Select Case (naCnt)
                        Case 1
                            idx_NA1 = M
                        Case 2
                            idx_NA2 = M
                        Case 3
                            idx_NA3 = M
                        Case 4
                            idx_NA4 = M
                        Case Else
                        End Select
                    ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                        idx_USI_LSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                        idx_USI_MSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                        idx_USO_LSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                        idx_USO_MSBbitCycle = M
                    ElseIf (offcolStr Like UCase("program*stage")) Then
                        idx_Stage = M
                    ElseIf (offcolStr Like UCase("Low*Limit")) Then
                        idx_LoLMT = M
                    ElseIf (offcolStr Like UCase("High*Limit")) Then
                        idx_HiLMT = M
                    ElseIf (offcolStr Like UCase("*Resolution")) Then
                        idx_Resolution = M
                    ElseIf (offcolStr Like UCase("Algorithm")) Then
                        idx_Algorithm = M
                    ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                        idx_Comment = M
                    ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                        idx_DefaultOrReal = M
                    ElseIf (offcolStr Like UCase("Default*Value")) Then
                        idx_DefaultValue = M
                    ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                        idx_Difference = M
                        findAllParamHeader = True
                        Exit Do
                    End If
                Loop While (findAllParamHeader = False)
                ''''-----------------------------------------------------------------------------------------------------------
                GoTo nextRow
            Else
                ''''case "bank_UDR_E eFuse Bit Def"
                ReDim UDRE_Fuse.Category(m_CateNumSize) ''''initialize
                n = 0
                findECIDHeader = False
                findCFGHeader = False
                findUIDHeader = False
                findUDRHeader = False
                findSENHeader = False
                findMONHeader = False
                findCMPHeader = False
                findUDRE_Header = True
                findUDRP_Header = False
                findCMPE_Header = False
                findCMPP_Header = False
    
                ''''-----------------------------------------------------------------------------------------------------------
                ''''Get the Specific Index
                M = 0: naCnt = 0
                offcolStr = ""
                findAllParamHeader = False
    
                ''''Initialize
                idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
                idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
                idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
                idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
                idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
                idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1
    
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    offcolStr = UCase(Trim(offCell.Value))
                    DebugPrintLog "4... UDRE input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                    
                    If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                        idx_LSBbit = M
                    ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                        idx_MSBbit = M
                    ElseIf (offcolStr Like UCase("Bit*Width")) Then
                        idx_BitWidth = M
                    ElseIf (offcolStr = UCase("N/A")) Then
                        naCnt = naCnt + 1
                        Select Case (naCnt)
                        Case 1
                            idx_NA1 = M
                        Case 2
                            idx_NA2 = M
                        Case 3
                            idx_NA3 = M
                        Case 4
                            idx_NA4 = M
                        Case Else
                        End Select
                    ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                        idx_USI_LSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                        idx_USI_MSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                        idx_USO_LSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                        idx_USO_MSBbitCycle = M
                    ElseIf (offcolStr Like UCase("program*stage")) Then
                        idx_Stage = M
                    ElseIf (offcolStr Like UCase("Low*Limit")) Then
                        idx_LoLMT = M
                    ElseIf (offcolStr Like UCase("High*Limit")) Then
                        idx_HiLMT = M
                    ElseIf (offcolStr Like UCase("*Resolution")) Then
                        idx_Resolution = M
                    ElseIf (offcolStr Like UCase("Algorithm")) Then
                        idx_Algorithm = M
                    ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                        idx_Comment = M
                    ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                        idx_DefaultOrReal = M
                    ElseIf (offcolStr Like UCase("Default*Value")) Then
                        idx_DefaultValue = M
                    ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                        idx_Difference = M
                        findAllParamHeader = True
                        Exit Do
                    End If
                Loop While (findAllParamHeader = False)
                ''''-----------------------------------------------------------------------------------------------------------
                GoTo nextRow
            End If

        ''''20171103 add
        ElseIf (m_cellStr Like UCase("*UDR_P*Bit*Def*")) Then

            ''''20180104 update per customer format request
            If (m_cellStr Like UCase("*CMP*Bit*Def*")) Then ''''case "bank_UDR_P CMP eFuse Bit Def"
                ReDim CMPP_Fuse.Category(m_CateNumSize) ''''initialize
                n = 0
                findECIDHeader = False
                findCFGHeader = False
                findUIDHeader = False
                findUDRHeader = False
                findSENHeader = False
                findMONHeader = False
                findCMPHeader = False
                findUDRE_Header = False
                findUDRP_Header = False
                findCMPE_Header = False
                findCMPP_Header = True
                ''''-----------------------------------------------------------------------------------------------------------
                ''''Get the Specific Index
                M = 0: naCnt = 0
                offcolStr = ""
                findAllParamHeader = False
    
                ''''Initialize
                idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
                idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
                idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
                idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
                idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
                idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1
    
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    offcolStr = UCase(Trim(offCell.Value))
                    DebugPrintLog "5... CMPE input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                    
                    If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                        idx_LSBbit = M
                    ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                        idx_MSBbit = M
                    ElseIf (offcolStr Like UCase("Bit*Width")) Then
                        idx_BitWidth = M
                    ElseIf (offcolStr = UCase("N/A")) Then
                        naCnt = naCnt + 1
                        Select Case (naCnt)
                        Case 1
                            idx_NA1 = M
                        Case 2
                            idx_NA2 = M
                        Case 3
                            idx_NA3 = M
                        Case 4
                            idx_NA4 = M
                        Case Else
                        End Select
                    ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                        idx_USI_LSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                        idx_USI_MSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                        idx_USO_LSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                        idx_USO_MSBbitCycle = M
                    ElseIf (offcolStr Like UCase("program*stage")) Then
                        idx_Stage = M
                    ElseIf (offcolStr Like UCase("Low*Limit")) Then
                        idx_LoLMT = M
                    ElseIf (offcolStr Like UCase("High*Limit")) Then
                        idx_HiLMT = M
                    ElseIf (offcolStr Like UCase("*Resolution")) Then
                        idx_Resolution = M
                    ElseIf (offcolStr Like UCase("Algorithm")) Then
                        idx_Algorithm = M
                    ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                        idx_Comment = M
                    ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                        idx_DefaultOrReal = M
                    ElseIf (offcolStr Like UCase("Default*Value")) Then
                        idx_DefaultValue = M
                    ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                        idx_Difference = M
                        findAllParamHeader = True
                        Exit Do
                    End If
                Loop While (findAllParamHeader = False)
                ''''-----------------------------------------------------------------------------------------------------------
                GoTo nextRow
            Else
                ''''case "bank_UDR_P eFuse Bit Def"
                ReDim UDRP_Fuse.Category(m_CateNumSize) ''''initialize
                n = 0
                findECIDHeader = False
                findCFGHeader = False
                findUIDHeader = False
                findUDRHeader = False
                findSENHeader = False
                findMONHeader = False
                findCMPHeader = False
                findUDRE_Header = False
                findUDRP_Header = True
                findCMPE_Header = False
                findCMPP_Header = False
    
                ''''-----------------------------------------------------------------------------------------------------------
                ''''Get the Specific Index
                M = 0: naCnt = 0
                offcolStr = ""
                findAllParamHeader = False
    
                ''''Initialize
                idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
                idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
                idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
                idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
                idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
                idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1
    
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    offcolStr = UCase(Trim(offCell.Value))
                    DebugPrintLog "4... UDRE input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                    
                    If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                        idx_LSBbit = M
                    ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                        idx_MSBbit = M
                    ElseIf (offcolStr Like UCase("Bit*Width")) Then
                        idx_BitWidth = M
                    ElseIf (offcolStr = UCase("N/A")) Then
                        naCnt = naCnt + 1
                        Select Case (naCnt)
                        Case 1
                            idx_NA1 = M
                        Case 2
                            idx_NA2 = M
                        Case 3
                            idx_NA3 = M
                        Case 4
                            idx_NA4 = M
                        Case Else
                        End Select
                    ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                        idx_USI_LSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                        idx_USI_MSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                        idx_USO_LSBbitCycle = M
                    ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                        idx_USO_MSBbitCycle = M
                    ElseIf (offcolStr Like UCase("program*stage")) Then
                        idx_Stage = M
                    ElseIf (offcolStr Like UCase("Low*Limit")) Then
                        idx_LoLMT = M
                    ElseIf (offcolStr Like UCase("High*Limit")) Then
                        idx_HiLMT = M
                    ElseIf (offcolStr Like UCase("*Resolution")) Then
                        idx_Resolution = M
                    ElseIf (offcolStr Like UCase("Algorithm")) Then
                        idx_Algorithm = M
                    ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                        idx_Comment = M
                    ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                        idx_DefaultOrReal = M
                    ElseIf (offcolStr Like UCase("Default*Value")) Then
                        idx_DefaultValue = M
                    ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                        idx_Difference = M
                        findAllParamHeader = True
                        Exit Do
                    End If
                Loop While (findAllParamHeader = False)
                ''''-----------------------------------------------------------------------------------------------------------
                GoTo nextRow
            End If
        ElseIf (m_cellStr Like UCase("*UDR*Bit*Def*")) Then

            ReDim UDRFuse.Category(m_CateNumSize) ''''initialize
            n = 0
            findECIDHeader = False
            findCFGHeader = False
            findUIDHeader = False
            findUDRHeader = True
            findSENHeader = False
            findMONHeader = False
            findCMPHeader = False
            findUDRE_Header = False
            findUDRP_Header = False
            findCMPE_Header = False
            findCMPP_Header = False

            ''''-----------------------------------------------------------------------------------------------------------
            ''''Get the Specific Index
            M = 0: naCnt = 0
            offcolStr = ""
            findAllParamHeader = False

            ''''Initialize
            idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
            idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
            idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
            idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
            idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
            idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1

            Do
                M = M + 1
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                offcolStr = UCase(Trim(offCell.Value))
                DebugPrintLog "4... UDR input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                
                If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                    idx_LSBbit = M
                ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                    idx_MSBbit = M
                ElseIf (offcolStr Like UCase("Bit*Width")) Then
                    idx_BitWidth = M
                ElseIf (offcolStr = UCase("N/A")) Then
                    naCnt = naCnt + 1
                    Select Case (naCnt)
                    Case 1
                        idx_NA1 = M
                    Case 2
                        idx_NA2 = M
                    Case 3
                        idx_NA3 = M
                    Case 4
                        idx_NA4 = M
                    Case Else
                    End Select
                ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                    idx_USI_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                    idx_USI_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                    idx_USO_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                    idx_USO_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("program*stage")) Then
                    idx_Stage = M
                ElseIf (offcolStr Like UCase("Low*Limit")) Then
                    idx_LoLMT = M
                ElseIf (offcolStr Like UCase("High*Limit")) Then
                    idx_HiLMT = M
                ElseIf (offcolStr Like UCase("*Resolution")) Then
                    idx_Resolution = M
                ElseIf (offcolStr Like UCase("Algorithm")) Then
                    idx_Algorithm = M
                ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                    idx_Comment = M
                ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                    idx_DefaultOrReal = M
                ElseIf (offcolStr Like UCase("Default*Value")) Then
                    idx_DefaultValue = M
                ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                    idx_Difference = M
                    findAllParamHeader = True
                    Exit Do
                End If
            Loop While (findAllParamHeader = False)
            ''''-----------------------------------------------------------------------------------------------------------
            GoTo nextRow

        ElseIf (m_cellStr Like UCase("*Sen*Bit*Def*")) Then ''''was m_cellStr Like UCase("Sensor*Bit*Def*")

            ReDim SENFuse.Category(m_CateNumSize) ''''initialize
            n = 0
            findECIDHeader = False
            findCFGHeader = False
            findUIDHeader = False
            findUDRHeader = False
            findSENHeader = True
            findMONHeader = False
            findCMPHeader = False
            findUDRE_Header = False
            findUDRP_Header = False
            findCMPE_Header = False
            findCMPP_Header = False
            
            ''''-----------------------------------------------------------------------------------------------------------
            ''''Get the Specific Index
            M = 0: naCnt = 0
            offcolStr = ""
            findAllParamHeader = False

            ''''Initialize
            idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
            idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
            idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
            idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
            idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
            idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1

            Do
                M = M + 1
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                offcolStr = UCase(Trim(offCell.Value))
                DebugPrintLog "5... SEN input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                
                If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                    idx_LSBbit = M
                ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                    idx_MSBbit = M
                ElseIf (offcolStr Like UCase("Bit*Width")) Then
                    idx_BitWidth = M
                ElseIf (offcolStr = UCase("N/A")) Then
                    naCnt = naCnt + 1
                    Select Case (naCnt)
                    Case 1
                        idx_NA1 = M
                    Case 2
                        idx_NA2 = M
                    Case 3
                        idx_NA3 = M
                    Case 4
                        idx_NA4 = M
                    Case Else
                    End Select
                ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                    idx_USI_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                    idx_USI_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                    idx_USO_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                    idx_USO_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("program*stage")) Then
                    idx_Stage = M
                ElseIf (offcolStr Like UCase("Low*Limit")) Then
                    idx_LoLMT = M
                ElseIf (offcolStr Like UCase("High*Limit")) Then
                    idx_HiLMT = M
                ElseIf (offcolStr Like UCase("*Resolution")) Then
                    idx_Resolution = M
                ElseIf (offcolStr Like UCase("Algorithm")) Then
                    idx_Algorithm = M
                ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                    idx_Comment = M
                ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                    idx_DefaultOrReal = M
                ElseIf (offcolStr Like UCase("Default*Value")) Then
                    idx_DefaultValue = M
                ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                    idx_Difference = M
                    findAllParamHeader = True
                    Exit Do
                End If
            Loop While (findAllParamHeader = False)
            ''''-----------------------------------------------------------------------------------------------------------
            GoTo nextRow
        
        ElseIf (m_cellStr Like UCase("*MON*Bit*Def*")) Then ''''was m_cellStr Like UCase("*MONITOR*Bit*Def*")

            ReDim MONFuse.Category(m_CateNumSize) ''''initialize
            n = 0
            findECIDHeader = False
            findCFGHeader = False
            findUIDHeader = False
            findUDRHeader = False
            findSENHeader = False
            findMONHeader = True
            findCMPHeader = False
            findUDRE_Header = False
            findUDRP_Header = False
            findCMPE_Header = False
            findCMPP_Header = False
            ''''-----------------------------------------------------------------------------------------------------------
            ''''Get the Specific Index
            M = 0: naCnt = 0
            offcolStr = ""
            findAllParamHeader = False

            ''''Initialize
            idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
            idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
            idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
            idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
            idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
            idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1

            Do
                M = M + 1
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                offcolStr = UCase(Trim(offCell.Value))
                DebugPrintLog "5... MON input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                
                If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                    idx_LSBbit = M
                ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                    idx_MSBbit = M
                ElseIf (offcolStr Like UCase("Bit*Width")) Then
                    idx_BitWidth = M
                ElseIf (offcolStr = UCase("N/A")) Then
                    naCnt = naCnt + 1
                    Select Case (naCnt)
                    Case 1
                        idx_NA1 = M
                    Case 2
                        idx_NA2 = M
                    Case 3
                        idx_NA3 = M
                    Case 4
                        idx_NA4 = M
                    Case Else
                    End Select
                ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                    idx_USI_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                    idx_USI_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                    idx_USO_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                    idx_USO_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("program*stage")) Then
                    idx_Stage = M
                ElseIf (offcolStr Like UCase("Low*Limit")) Then
                    idx_LoLMT = M
                ElseIf (offcolStr Like UCase("High*Limit")) Then
                    idx_HiLMT = M
                ElseIf (offcolStr Like UCase("*Resolution")) Then
                    idx_Resolution = M
                ElseIf (offcolStr Like UCase("Algorithm")) Then
                    idx_Algorithm = M
                ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                    idx_Comment = M
                ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                    idx_DefaultOrReal = M
                ElseIf (offcolStr Like UCase("Default*Value")) Then
                    idx_DefaultValue = M
                ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                    idx_Difference = M
                    findAllParamHeader = True
                    Exit Do
                End If
            Loop While (findAllParamHeader = False)
            ''''-----------------------------------------------------------------------------------------------------------
            GoTo nextRow

        ''''20171103 add
        ElseIf (m_cellStr Like UCase("*CMP_E*Bit*Def*")) Then
            ReDim CMPE_Fuse.Category(m_CateNumSize) ''''initialize
            n = 0
            findECIDHeader = False
            findCFGHeader = False
            findUIDHeader = False
            findUDRHeader = False
            findSENHeader = False
            findMONHeader = False
            findCMPHeader = False
            findUDRE_Header = False
            findUDRP_Header = False
            findCMPE_Header = True
            findCMPP_Header = False
            ''''-----------------------------------------------------------------------------------------------------------
            ''''Get the Specific Index
            M = 0: naCnt = 0
            offcolStr = ""
            findAllParamHeader = False

            ''''Initialize
            idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
            idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
            idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
            idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
            idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
            idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1

            Do
                M = M + 1
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                offcolStr = UCase(Trim(offCell.Value))
                DebugPrintLog "5... CMPE input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                
                If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                    idx_LSBbit = M
                ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                    idx_MSBbit = M
                ElseIf (offcolStr Like UCase("Bit*Width")) Then
                    idx_BitWidth = M
                ElseIf (offcolStr = UCase("N/A")) Then
                    naCnt = naCnt + 1
                    Select Case (naCnt)
                    Case 1
                        idx_NA1 = M
                    Case 2
                        idx_NA2 = M
                    Case 3
                        idx_NA3 = M
                    Case 4
                        idx_NA4 = M
                    Case Else
                    End Select
                ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                    idx_USI_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                    idx_USI_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                    idx_USO_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                    idx_USO_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("program*stage")) Then
                    idx_Stage = M
                ElseIf (offcolStr Like UCase("Low*Limit")) Then
                    idx_LoLMT = M
                ElseIf (offcolStr Like UCase("High*Limit")) Then
                    idx_HiLMT = M
                ElseIf (offcolStr Like UCase("*Resolution")) Then
                    idx_Resolution = M
                ElseIf (offcolStr Like UCase("Algorithm")) Then
                    idx_Algorithm = M
                ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                    idx_Comment = M
                ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                    idx_DefaultOrReal = M
                ElseIf (offcolStr Like UCase("Default*Value")) Then
                    idx_DefaultValue = M
                ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                    idx_Difference = M
                    findAllParamHeader = True
                    Exit Do
                End If
            Loop While (findAllParamHeader = False)
            ''''-----------------------------------------------------------------------------------------------------------
            GoTo nextRow

        ''''20171103 add
        ElseIf (m_cellStr Like UCase("*CMP_P*Bit*Def*")) Then
            ReDim CMPP_Fuse.Category(m_CateNumSize) ''''initialize
            n = 0
            findECIDHeader = False
            findCFGHeader = False
            findUIDHeader = False
            findUDRHeader = False
            findSENHeader = False
            findMONHeader = False
            findCMPHeader = False
            findUDRE_Header = False
            findUDRP_Header = False
            findCMPE_Header = False
            findCMPP_Header = True
            ''''-----------------------------------------------------------------------------------------------------------
            ''''Get the Specific Index
            M = 0: naCnt = 0
            offcolStr = ""
            findAllParamHeader = False

            ''''Initialize
            idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
            idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
            idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
            idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
            idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
            idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1

            Do
                M = M + 1
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                offcolStr = UCase(Trim(offCell.Value))
                DebugPrintLog "5... CMPE input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                
                If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                    idx_LSBbit = M
                ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                    idx_MSBbit = M
                ElseIf (offcolStr Like UCase("Bit*Width")) Then
                    idx_BitWidth = M
                ElseIf (offcolStr = UCase("N/A")) Then
                    naCnt = naCnt + 1
                    Select Case (naCnt)
                    Case 1
                        idx_NA1 = M
                    Case 2
                        idx_NA2 = M
                    Case 3
                        idx_NA3 = M
                    Case 4
                        idx_NA4 = M
                    Case Else
                    End Select
                ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                    idx_USI_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                    idx_USI_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                    idx_USO_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                    idx_USO_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("program*stage")) Then
                    idx_Stage = M
                ElseIf (offcolStr Like UCase("Low*Limit")) Then
                    idx_LoLMT = M
                ElseIf (offcolStr Like UCase("High*Limit")) Then
                    idx_HiLMT = M
                ElseIf (offcolStr Like UCase("*Resolution")) Then
                    idx_Resolution = M
                ElseIf (offcolStr Like UCase("Algorithm")) Then
                    idx_Algorithm = M
                ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                    idx_Comment = M
                ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                    idx_DefaultOrReal = M
                ElseIf (offcolStr Like UCase("Default*Value")) Then
                    idx_DefaultValue = M
                ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                    idx_Difference = M
                    findAllParamHeader = True
                    Exit Do
                End If
            Loop While (findAllParamHeader = False)
            ''''-----------------------------------------------------------------------------------------------------------
            GoTo nextRow

        ElseIf (m_cellStr Like UCase("*CMP*Bit*Def*")) Then

            ReDim CMPFuse.Category(m_CateNumSize) ''''initialize
            n = 0
            findECIDHeader = False
            findCFGHeader = False
            findUIDHeader = False
            findUDRHeader = False
            findSENHeader = False
            findMONHeader = False
            findCMPHeader = True
            findUDRE_Header = False
            findUDRP_Header = False
            findCMPE_Header = False
            findCMPP_Header = False
            ''''-----------------------------------------------------------------------------------------------------------
            ''''Get the Specific Index
            M = 0: naCnt = 0
            offcolStr = ""
            findAllParamHeader = False

            ''''Initialize
            idx_seqSTART = -1: idx_seqEnd = -1: idx_MSBbit = -1: idx_LSBbit = -1: idx_BitWidth = -1
            idx_NA1 = -1: idx_NA2 = -1: idx_NA3 = -1: idx_NA4 = -1
            idx_USI_MSBbitCycle = -1: idx_USI_LSBbitCycle = -1: idx_USO_MSBbitCycle = -1: idx_USO_LSBbitCycle = -1
            idx_Stage = -1: idx_LoLMT = -1: idx_HiLMT = -1
            idx_Resolution = -1: idx_Algorithm = -1: idx_Comment = -1
            idx_DefaultOrReal = -1: idx_DefaultValue = -1: idx_Difference = -1

            Do
                M = M + 1
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                offcolStr = UCase(Trim(offCell.Value))
                DebugPrintLog "5... CMP input offcolStr=" + offcolStr + ", Offset Index=" + CStr(M)
                
                If (offcolStr Like UCase("LSB*Bit") Or offcolStr Like UCase("Seq*Start")) Then
                    idx_LSBbit = M
                ElseIf (offcolStr Like UCase("MSB*Bit") Or offcolStr Like UCase("Seq*End")) Then
                    idx_MSBbit = M
                ElseIf (offcolStr Like UCase("Bit*Width")) Then
                    idx_BitWidth = M
                ElseIf (offcolStr = UCase("N/A")) Then
                    naCnt = naCnt + 1
                    Select Case (naCnt)
                    Case 1
                        idx_NA1 = M
                    Case 2
                        idx_NA2 = M
                    Case 3
                        idx_NA3 = M
                    Case 4
                        idx_NA4 = M
                    Case Else
                    End Select
                ElseIf (offcolStr Like UCase("USI*LSB*Bit*Cycle")) Then
                    idx_USI_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USI*MSB*Bit*Cycle")) Then
                    idx_USI_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*LSB*Bit*Cycle")) Then
                    idx_USO_LSBbitCycle = M
                ElseIf (offcolStr Like UCase("USO*MSB*Bit*Cycle")) Then
                    idx_USO_MSBbitCycle = M
                ElseIf (offcolStr Like UCase("program*stage")) Then
                    idx_Stage = M
                ElseIf (offcolStr Like UCase("Low*Limit")) Then
                    idx_LoLMT = M
                ElseIf (offcolStr Like UCase("High*Limit")) Then
                    idx_HiLMT = M
                ElseIf (offcolStr Like UCase("*Resolution")) Then
                    idx_Resolution = M
                ElseIf (offcolStr Like UCase("Algorithm")) Then
                    idx_Algorithm = M
                ElseIf (offcolStr Like UCase("*Comment*")) Or (offcolStr Like UCase("*Description*")) Then
                    idx_Comment = M
                ElseIf (offcolStr Like UCase("*Default*or*Real*")) Then
                    idx_DefaultOrReal = M
                ElseIf (offcolStr Like UCase("Default*Value")) Then
                    idx_DefaultValue = M
                ElseIf (offcolStr Like UCase("Difference")) Then ''Must be the last
                    idx_Difference = M
                    findAllParamHeader = True
                    Exit Do
                End If
            Loop While (findAllParamHeader = False)
            ''''-----------------------------------------------------------------------------------------------------------
            GoTo nextRow

        End If

        If (m_cellStr <> "") Then
            
            If (findECIDHeader) Then

                ECIDFuse.Category(n).index = n
                ECIDFuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_ECID = CateCnt_ECID + 1
                DebugPrintLog "1...ECIDFuse.Category Index=" + CStr(n) + ":: " + ECIDFuse.Category(n).Name
                DebugPrintLog "1...ECIDFuse.Category CateCnt_ECID=" + CStr(CateCnt_ECID)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or UCase(offcolStr) Like "*REAL*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_MSBbit ''''<NOTICE> idx_seqSTART
                        ECIDFuse.Category(n).MSBbit = CLng(offcolStr)
                        ECIDFuse.Category(n).SeqStart = ECIDFuse.Category(n).MSBbit
                    Case idx_LSBbit ''''<NOTICE> idx_seqEnd
                        ECIDFuse.Category(n).LSBbit = CLng(offcolStr)
                        ECIDFuse.Category(n).SeqEnd = ECIDFuse.Category(n).LSBbit
                    Case idx_BitWidth
                        ECIDFuse.Category(n).BitWidth = CLng(offcolStr)
                    Case idx_Stage
                        ECIDFuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        If UCase(ECIDFuse.Category(n).Name) Like "*X_COORD*" Or UCase(ECIDFuse.Category(n).Name) Like "*Y_COORD*" Or UCase(ECIDFuse.Category(n).Name) Like "*WAFER_ID*" Then ''' 20180316
                            If UCase(offcolStr) Like UCase("x*") Then
                                offcolStr = Replace(UCase(offcolStr), "X", "", 1)
                                offcolStr = CLng("&H" & CStr(offcolStr))
                            End If
                        End If
                        ECIDFuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        If UCase(ECIDFuse.Category(n).Name) Like "*X_COORD*" Or UCase(ECIDFuse.Category(n).Name) Like "*Y_COORD*" Or UCase(ECIDFuse.Category(n).Name) Like "*WAFER_ID*" Then ''' 20180316
                            If UCase(offcolStr) Like UCase("x*") Then
                                offcolStr = Replace(UCase(offcolStr), "X", "", 1)
                                offcolStr = CLng("&H" & CStr(offcolStr))
                            End If
                        End If
                        ECIDFuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        ECIDFuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        ECIDFuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        ECIDFuse.Category(n).comment = offcolStr
                        DebugPrintLog "    1sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        ECIDFuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    1sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        ECIDFuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    1sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    1sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select
                    
                    DebugPrintLog "    1   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing
    
                n = n + 1
            
            ElseIf (findCFGHeader) Then
            
                CFGFuse.Category(n).index = n
                CFGFuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_CFG = CateCnt_CFG + 1
                DebugPrintLog "2...CFGFuse.Category Index=" + CStr(n) + ":: " + CFGFuse.Category(n).Name
                DebugPrintLog "2...CFGFuse.Category CateCnt_CFG=" + CStr(CateCnt_CFG)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or _
                            UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or _
                            UCase(offcolStr) Like "*REAL*" Or UCase(offcolStr) Like "*BINCUT*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_LSBbit
                        CFGFuse.Category(n).LSBbit = CLng(offcolStr)
                        CFGFuse.Category(n).SeqStart = CFGFuse.Category(n).LSBbit
                    Case idx_MSBbit
                        CFGFuse.Category(n).MSBbit = CLng(offcolStr)
                        CFGFuse.Category(n).SeqEnd = CFGFuse.Category(n).MSBbit
                    Case idx_BitWidth
                        CFGFuse.Category(n).BitWidth = CLng(offcolStr)
                    Case idx_Stage
                        CFGFuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        CFGFuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        CFGFuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        CFGFuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        CFGFuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        CFGFuse.Category(n).comment = offcolStr
                        DebugPrintLog "    2sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        CFGFuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    2sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        CFGFuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    2sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    2sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select
                    
                    DebugPrintLog "    2   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing
    
                n = n + 1

            ElseIf (findUIDHeader) Then
                ''ByPass and do-nothing
                UIDFuse.Category(n).index = n
                UIDFuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_UID = CateCnt_UID + 1
                DebugPrintLog "3...UIDFuse.Category Index=" + CStr(n) + ":: " + UIDFuse.Category(n).Name
                DebugPrintLog "3...UIDFuse.Category CateCnt_UID=" + CStr(CateCnt_UID)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or UCase(offcolStr) Like "*REAL*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_LSBbit
                        UIDFuse.Category(n).LSBbit = CLng(offcolStr)
                        UIDFuse.Category(n).SeqStart = UIDFuse.Category(n).LSBbit
                    Case idx_MSBbit
                        UIDFuse.Category(n).MSBbit = CLng(offcolStr)
                        UIDFuse.Category(n).SeqEnd = UIDFuse.Category(n).MSBbit
                    Case idx_BitWidth
                        UIDFuse.Category(n).BitWidth = CLng(offcolStr)
                    Case idx_Stage
                        UIDFuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        UIDFuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        UIDFuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        UIDFuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        UIDFuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        UIDFuse.Category(n).comment = offcolStr
                        DebugPrintLog "    3sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        UIDFuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    3sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        UIDFuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    3sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    3sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select
                    
                    DebugPrintLog "    3   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing

                n = n + 1

            ElseIf (findUDRHeader) Then
            
                UDRFuse.Category(n).index = n
                UDRFuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_UDR = CateCnt_UDR + 1
                DebugPrintLog "4...UDRFuse.Category Index=" + CStr(n) + ":: " + UDRFuse.Category(n).Name
                DebugPrintLog "4...UDRFuse.Category CateCnt_UDR=" + CStr(CateCnt_UDR)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or _
                            UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or _
                            UCase(offcolStr) Like "*REAL*" Or UCase(offcolStr) Like "*BINCUT*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_LSBbit
                        UDRFuse.Category(n).LSBbit = CLng(offcolStr)
                        UDRFuse.Category(n).SeqStart = UDRFuse.Category(n).LSBbit
                    Case idx_MSBbit
                        UDRFuse.Category(n).MSBbit = CLng(offcolStr)
                        UDRFuse.Category(n).SeqEnd = UDRFuse.Category(n).MSBbit
                    Case idx_BitWidth
                        UDRFuse.Category(n).BitWidth = CLng(offcolStr)
                    Case idx_USI_LSBbitCycle
                        UDRFuse.Category(n).USILSBCycle = CLng(offcolStr)
                    Case idx_USI_MSBbitCycle
                        UDRFuse.Category(n).USIMSBCycle = CLng(offcolStr)
                    Case idx_USO_LSBbitCycle
                        UDRFuse.Category(n).USOLSBCycle = CLng(offcolStr)
                    Case idx_USO_MSBbitCycle
                        UDRFuse.Category(n).USOMSBCycle = CLng(offcolStr)
                    Case idx_Stage
                        UDRFuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        UDRFuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        UDRFuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        UDRFuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        UDRFuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        UDRFuse.Category(n).comment = offcolStr
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        UDRFuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        UDRFuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select
                    
                    DebugPrintLog "    4   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing

                n = n + 1

            ElseIf (findSENHeader) Then
            
                SENFuse.Category(n).index = n
                SENFuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_SEN = CateCnt_SEN + 1
                DebugPrintLog "5...SENFuse.Category Index=" + CStr(n) + ":: " + SENFuse.Category(n).Name
                DebugPrintLog "5...SENFuse.Category CateCnt_SEN=" + CStr(CateCnt_SEN)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or UCase(offcolStr) Like "*REAL*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_LSBbit
                        SENFuse.Category(n).LSBbit = CInt(offcolStr)
                        SENFuse.Category(n).SeqStart = SENFuse.Category(n).LSBbit
                    Case idx_MSBbit
                        SENFuse.Category(n).MSBbit = CInt(offcolStr)
                        SENFuse.Category(n).SeqEnd = SENFuse.Category(n).MSBbit
                    Case idx_BitWidth
                        SENFuse.Category(n).BitWidth = CInt(offcolStr)
                    Case idx_Stage
                        SENFuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        SENFuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        SENFuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        SENFuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        SENFuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        SENFuse.Category(n).comment = offcolStr
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        SENFuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        SENFuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select

                    DebugPrintLog "    5   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing

                n = n + 1
                
            ElseIf (findMONHeader) Then
            
                MONFuse.Category(n).index = n
                MONFuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_MON = CateCnt_MON + 1
                DebugPrintLog "5...MONFuse.Category Index=" + CStr(n) + ":: " + MONFuse.Category(n).Name
                DebugPrintLog "5...MONFuse.Category CateCnt_MON=" + CStr(CateCnt_MON)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or UCase(offcolStr) Like "*REAL*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_LSBbit
                        MONFuse.Category(n).LSBbit = CInt(offcolStr)
                        MONFuse.Category(n).SeqStart = MONFuse.Category(n).LSBbit
                    Case idx_MSBbit
                        MONFuse.Category(n).MSBbit = CInt(offcolStr)
                        MONFuse.Category(n).SeqEnd = MONFuse.Category(n).MSBbit
                    Case idx_BitWidth
                        MONFuse.Category(n).BitWidth = CInt(offcolStr)
                    Case idx_Stage
                        MONFuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        MONFuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        MONFuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        MONFuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        MONFuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        MONFuse.Category(n).comment = offcolStr
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        MONFuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        MONFuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select

                    DebugPrintLog "    5   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing

                n = n + 1

            ElseIf (findCMPHeader) Then
            
                CMPFuse.Category(n).index = n
                CMPFuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_CMP = CateCnt_CMP + 1
                DebugPrintLog "6...CMPFuse.Category Index=" + CStr(n) + ":: " + CMPFuse.Category(n).Name
                DebugPrintLog "6...CMPFuse.Category CateCnt_CMP=" + CStr(CateCnt_CMP)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or UCase(offcolStr) Like "*REAL*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_LSBbit
                        CMPFuse.Category(n).LSBbit = CInt(offcolStr)
                        CMPFuse.Category(n).SeqStart = CMPFuse.Category(n).LSBbit
                    Case idx_MSBbit
                        CMPFuse.Category(n).MSBbit = CInt(offcolStr)
                        CMPFuse.Category(n).SeqEnd = CMPFuse.Category(n).MSBbit
                    Case idx_BitWidth
                        CMPFuse.Category(n).BitWidth = CInt(offcolStr)
                    Case idx_USI_LSBbitCycle
                        CMPFuse.Category(n).USILSBCycle = CLng(offcolStr)
                    Case idx_USI_MSBbitCycle
                        CMPFuse.Category(n).USIMSBCycle = CLng(offcolStr)
                    Case idx_USO_LSBbitCycle
                        CMPFuse.Category(n).USOLSBCycle = CLng(offcolStr)
                    Case idx_USO_MSBbitCycle
                        CMPFuse.Category(n).USOMSBCycle = CLng(offcolStr)
                    Case idx_Stage
                        CMPFuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        CMPFuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        CMPFuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        CMPFuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        CMPFuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        CMPFuse.Category(n).comment = offcolStr
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        CMPFuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        CMPFuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    5sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select

                    DebugPrintLog "    6   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing

                n = n + 1

            ElseIf (findUDRE_Header) Then
            
                UDRE_Fuse.Category(n).index = n
                UDRE_Fuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_UDRE = CateCnt_UDRE + 1
                DebugPrintLog "4...UDRE_Fuse.Category Index=" + CStr(n) + ":: " + UDRE_Fuse.Category(n).Name
                DebugPrintLog "4...UDRE_Fuse.Category CateCnt_UDRE=" + CStr(CateCnt_UDRE)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or _
                            UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or _
                            UCase(offcolStr) Like "*REAL*" Or UCase(offcolStr) Like "*BINCUT*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_LSBbit
                        UDRE_Fuse.Category(n).LSBbit = CLng(offcolStr)
                        UDRE_Fuse.Category(n).SeqStart = UDRE_Fuse.Category(n).LSBbit
                    Case idx_MSBbit
                        UDRE_Fuse.Category(n).MSBbit = CLng(offcolStr)
                        UDRE_Fuse.Category(n).SeqEnd = UDRE_Fuse.Category(n).MSBbit
                    Case idx_BitWidth
                        UDRE_Fuse.Category(n).BitWidth = CLng(offcolStr)
                    Case idx_USI_LSBbitCycle
                        UDRE_Fuse.Category(n).USILSBCycle = CLng(offcolStr)
                    Case idx_USI_MSBbitCycle
                        UDRE_Fuse.Category(n).USIMSBCycle = CLng(offcolStr)
                    Case idx_USO_LSBbitCycle
                        UDRE_Fuse.Category(n).USOLSBCycle = CLng(offcolStr)
                    Case idx_USO_MSBbitCycle
                        UDRE_Fuse.Category(n).USOMSBCycle = CLng(offcolStr)
                    Case idx_Stage
                        UDRE_Fuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        UDRE_Fuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        UDRE_Fuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        UDRE_Fuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        UDRE_Fuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        UDRE_Fuse.Category(n).comment = offcolStr
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        UDRE_Fuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        UDRE_Fuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select
                    
                    DebugPrintLog "    4   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing

                n = n + 1

            ElseIf (findUDRP_Header) Then
            
                UDRP_Fuse.Category(n).index = n
                UDRP_Fuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_UDRP = CateCnt_UDRP + 1
                DebugPrintLog "4...UDRP_Fuse.Category Index=" + CStr(n) + ":: " + UDRP_Fuse.Category(n).Name
                DebugPrintLog "4...UDRP_Fuse.Category CateCnt_UDRP=" + CStr(CateCnt_UDRP)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or _
                            UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or _
                            UCase(offcolStr) Like "*REAL*" Or UCase(offcolStr) Like "*BINCUT*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_LSBbit
                        UDRP_Fuse.Category(n).LSBbit = CLng(offcolStr)
                        UDRP_Fuse.Category(n).SeqStart = UDRP_Fuse.Category(n).LSBbit
                    Case idx_MSBbit
                        UDRP_Fuse.Category(n).MSBbit = CLng(offcolStr)
                        UDRP_Fuse.Category(n).SeqEnd = UDRP_Fuse.Category(n).MSBbit
                    Case idx_BitWidth
                        UDRP_Fuse.Category(n).BitWidth = CLng(offcolStr)
                    Case idx_USI_LSBbitCycle
                        UDRP_Fuse.Category(n).USILSBCycle = CLng(offcolStr)
                    Case idx_USI_MSBbitCycle
                        UDRP_Fuse.Category(n).USIMSBCycle = CLng(offcolStr)
                    Case idx_USO_LSBbitCycle
                        UDRP_Fuse.Category(n).USOLSBCycle = CLng(offcolStr)
                    Case idx_USO_MSBbitCycle
                        UDRP_Fuse.Category(n).USOMSBCycle = CLng(offcolStr)
                    Case idx_Stage
                        UDRP_Fuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        UDRP_Fuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        UDRP_Fuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        UDRP_Fuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        UDRP_Fuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        UDRP_Fuse.Category(n).comment = offcolStr
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        UDRP_Fuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        UDRP_Fuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select
                    
                    DebugPrintLog "    4   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing

                n = n + 1

            ElseIf (findCMPE_Header) Then
            
                CMPE_Fuse.Category(n).index = n
                CMPE_Fuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_CMPE = CateCnt_CMPE + 1
                DebugPrintLog "4...CMPE_Fuse.Category Index=" + CStr(n) + ":: " + CMPE_Fuse.Category(n).Name
                DebugPrintLog "4...CMPE_Fuse.Category CateCnt_CMPE=" + CStr(CateCnt_CMPE)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or UCase(offcolStr) Like "*REAL*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_LSBbit
                        CMPE_Fuse.Category(n).LSBbit = CLng(offcolStr)
                        CMPE_Fuse.Category(n).SeqStart = CMPE_Fuse.Category(n).LSBbit
                    Case idx_MSBbit
                        CMPE_Fuse.Category(n).MSBbit = CLng(offcolStr)
                        CMPE_Fuse.Category(n).SeqEnd = CMPE_Fuse.Category(n).MSBbit
                    Case idx_BitWidth
                        CMPE_Fuse.Category(n).BitWidth = CLng(offcolStr)
                    Case idx_USI_LSBbitCycle
                        CMPE_Fuse.Category(n).USILSBCycle = CLng(offcolStr)
                    Case idx_USI_MSBbitCycle
                        CMPE_Fuse.Category(n).USIMSBCycle = CLng(offcolStr)
                    Case idx_USO_LSBbitCycle
                        CMPE_Fuse.Category(n).USOLSBCycle = CLng(offcolStr)
                    Case idx_USO_MSBbitCycle
                        CMPE_Fuse.Category(n).USOMSBCycle = CLng(offcolStr)
                    Case idx_Stage
                        CMPE_Fuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        CMPE_Fuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        CMPE_Fuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        CMPE_Fuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        CMPE_Fuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        CMPE_Fuse.Category(n).comment = offcolStr
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        CMPE_Fuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        CMPE_Fuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select
                    
                    DebugPrintLog "    4   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing

                n = n + 1

            ElseIf (findCMPP_Header) Then
            
                CMPP_Fuse.Category(n).index = n
                CMPP_Fuse.Category(n).Name = CStr(Trim(myCell.Value))
                CateCnt_CMPP = CateCnt_CMPP + 1
                DebugPrintLog "4...CMPP_Fuse.Category Index=" + CStr(n) + ":: " + CMPP_Fuse.Category(n).Name
                DebugPrintLog "4...CMPP_Fuse.Category CateCnt_CMPP=" + CStr(CateCnt_CMPP)
                
                ''''Then get the following parameter per Category
                M = 0
                offcolStr = ""
                Do
                    M = M + 1
                    Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                    
                    offcolStr = Trim(offCell.Value)
                    If (offcolStr = "") Then
                        offcolStr = "0"
                        ''DebugPrintLog "offcolStr=" & CDbl(offcolStr)
                    ElseIf (M = idx_DefaultValue) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*" Or UCase(offcolStr) Like "*SEE*" Or UCase(offcolStr) Like "*REAL*") Then
                            offcolStr = "0"
                        End If
                    ElseIf (M = idx_Resolution) Then
                        If (UCase(offcolStr) = "N/A" Or UCase(offcolStr) = "NA" Or UCase(offcolStr) Like "ERR*") Then
                            offcolStr = "0"
                        End If
                    End If
                    
                    Select Case (M)
                    Case idx_LSBbit
                        CMPP_Fuse.Category(n).LSBbit = CLng(offcolStr)
                        CMPP_Fuse.Category(n).SeqStart = CMPP_Fuse.Category(n).LSBbit
                    Case idx_MSBbit
                        CMPP_Fuse.Category(n).MSBbit = CLng(offcolStr)
                        CMPP_Fuse.Category(n).SeqEnd = CMPP_Fuse.Category(n).MSBbit
                    Case idx_BitWidth
                        CMPP_Fuse.Category(n).BitWidth = CLng(offcolStr)
                    Case idx_USI_LSBbitCycle
                        CMPP_Fuse.Category(n).USILSBCycle = CLng(offcolStr)
                    Case idx_USI_MSBbitCycle
                        CMPP_Fuse.Category(n).USIMSBCycle = CLng(offcolStr)
                    Case idx_USO_LSBbitCycle
                        CMPP_Fuse.Category(n).USOLSBCycle = CLng(offcolStr)
                    Case idx_USO_MSBbitCycle
                        CMPP_Fuse.Category(n).USOMSBCycle = CLng(offcolStr)
                    Case idx_Stage
                        CMPP_Fuse.Category(n).Stage = offcolStr
                    Case idx_LoLMT
                        CMPP_Fuse.Category(n).LoLMT = CVar(offcolStr)
                    Case idx_HiLMT
                        CMPP_Fuse.Category(n).HiLMT = CVar(offcolStr)
                    Case idx_Resolution
                        CMPP_Fuse.Category(n).Resoultion = CDbl(offcolStr)
                    Case idx_Algorithm
                        CMPP_Fuse.Category(n).algorithm = offcolStr
                    Case idx_Comment
                        CMPP_Fuse.Category(n).comment = offcolStr
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (Comment/Description) = " + offcolStr
                    Case idx_DefaultOrReal
                        CMPP_Fuse.Category(n).Default_Real = offcolStr
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (DefaultOrReal) = " + offcolStr
                    Case idx_DefaultValue
                        CMPP_Fuse.Category(n).DefaultValue = CVar(offcolStr)
                    Case idx_NA1
                    Case idx_NA2
                    Case idx_NA3
                    Case idx_NA4
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (N/A) = " + offcolStr
                    Case idx_Difference
                        DebugPrintLog "    4sub...Parameter " + CStr(M) + " = (Difference) = " + offcolStr
                        Exit Do
                    Case Else
                        DebugPrintLog "Error on Select !!!"
                    End Select
                    
                    DebugPrintLog "    4   ...Parameter " + CStr(M) + " = " + offcolStr
                Loop While (M <= idx_Difference)
                ''''End of the Category Parsing

                n = n + 1

            End If
                        
        End If  ''''end of If(m_CellStr <> "") Then
    
nextRow:
        Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0) ''''search cell from Up to Down
        m_cellStr = UCase(Trim(myCell.Value))
        DebugPrintLog "0...input myCell=" + m_cellStr ''''CStr(myCell.Value)
    
    Loop ''''end of Do While (m_CellStr <> "END")
    
    ''''----------------------------------------------------------------------------------------------
    ''''After Sheet END
    ''''----------------------------------------------------------------------------------------------
    ''''Here is the case if the BDF sheet does NOT contain Anyone of ECID/CFG/UID/UDR/SEN/MON fuse content.
    ''''It is used to check the last one block,
    ''''because we can NOT expect the same order of the eFuse blocks through all projects.
    If (findECIDHeader) Then
        DebugPrintLog "1...ECIDFuse.Category Array Size = " + CStr(CateCnt_ECID)
        If (CateCnt_ECID > 0) Then
            gB_findECID_flag = True
            ReDim Preserve ECIDFuse.Category(CateCnt_ECID - 1) ''''NOT final dimension
        End If
        '''--------------------------------------------------------------------------
    End If
    
    If (findCFGHeader) Then
        DebugPrintLog "2...CFGFuse.Category Array Size = " + CStr(CateCnt_CFG)
        If (CateCnt_CFG > 0) Then
            gB_findCFG_flag = True
            ReDim Preserve CFGFuse.Category(CateCnt_CFG - 1) ''''final dimension
        End If
    End If
    
    If (findUIDHeader) Then
        DebugPrintLog "3......UIDFuse.Category Array Size = " + CStr(CateCnt_UID)
        If (CateCnt_UID > 0) Then
            gB_findUID_flag = True
            ReDim Preserve UIDFuse.Category(CateCnt_UID - 1) ''''final dimension
        End If
    End If

    If (findUDRHeader) Then
        DebugPrintLog "4...UDRFuse.Category Array Size = " + CStr(CateCnt_UDR)
        If (CateCnt_UDR > 0) Then
            gB_findUDR_flag = True
            ReDim Preserve UDRFuse.Category(CateCnt_UDR - 1) ''''final dimension
        End If
    End If
    
    If (findSENHeader) Then
        DebugPrintLog "5...SENFuse.Category Array Size = " + CStr(CateCnt_SEN)
        If (CateCnt_SEN > 0) Then
            gB_findSEN_flag = True
            ReDim Preserve SENFuse.Category(CateCnt_SEN - 1) ''''final dimension
        End If
    End If
        
    If (findMONHeader) Then
        DebugPrintLog "5...MONFuse.Category Array Size = " + CStr(CateCnt_MON)
        If (CateCnt_MON > 0) Then
            gB_findMON_flag = True
            ReDim Preserve MONFuse.Category(CateCnt_MON - 1) ''''final dimension
        End If
    End If
    
    If (findCMPHeader) Then
        DebugPrintLog "6...CMPFuse.Category Array Size = " + CStr(CateCnt_CMP)
        If (CateCnt_CMP > 0) Then
            gB_findCMP_flag = True
            ReDim Preserve CMPFuse.Category(CateCnt_CMP - 1) ''''final dimension
        End If
    End If

    If (findUDRE_Header) Then
        DebugPrintLog "7...UDRE_Fuse.Category Array Size = " + CStr(CateCnt_UDRE)
        If (CateCnt_UDRE > 0) Then
            gB_findUDRE_flag = True
            ReDim Preserve UDRE_Fuse.Category(CateCnt_UDRE - 1) ''''final dimension
        End If
    End If

    If (findUDRP_Header) Then
        DebugPrintLog "8...UDRP_Fuse.Category Array Size = " + CStr(CateCnt_UDRP)
        If (CateCnt_UDRP > 0) Then
            gB_findUDRP_flag = True
            ReDim Preserve UDRP_Fuse.Category(CateCnt_UDRP - 1) ''''final dimension
        End If
    End If

    If (findCMPE_Header) Then
        DebugPrintLog "9...CMPE_Fuse.Category Array Size = " + CStr(CateCnt_CMPE)
        If (CateCnt_CMPE > 0) Then
            gB_findCMPE_flag = True
            ReDim Preserve CMPE_Fuse.Category(CateCnt_CMPE - 1) ''''final dimension
        End If
    End If

    If (findCMPP_Header) Then
        DebugPrintLog "10...CMPP_Fuse.Category Array Size = " + CStr(CateCnt_CMPP)
        If (CateCnt_CMPP > 0) Then
            gB_findCMPP_flag = True
            ReDim Preserve CMPP_Fuse.Category(CateCnt_CMPP - 1) ''''final dimension
        End If
    End If

    findECIDHeader = False
    findCFGHeader = False
    findUIDHeader = False
    findUDRHeader = False
    findSENHeader = False
    findMONHeader = False
    findCMPHeader = False
    findUDRE_Header = False
    findUDRP_Header = False
    findCMPE_Header = False
    findCMPP_Header = False
    ''''----------------------------------------------------------------------------------------------
    
    ''''Final Summary
    If (gB_findECID_flag) Then DebugPrintLog "7...ECIDFuse.Category UBound = " + CStr(UBound(ECIDFuse.Category))
    If (gB_findCFG_flag) Then DebugPrintLog "7... CFGFuse.Category UBound = " + CStr(UBound(CFGFuse.Category))
    If (gB_findUID_flag) Then DebugPrintLog "7... UIDFuse.Category UBound = " + CStr(UBound(UIDFuse.Category))
    If (gB_findUDR_flag) Then DebugPrintLog "7... UDRFuse.Category UBound = " + CStr(UBound(UDRFuse.Category))
    If (gB_findSEN_flag) Then DebugPrintLog "7... SENFuse.Category UBound = " + CStr(UBound(SENFuse.Category))
    If (gB_findMON_flag) Then DebugPrintLog "7... MONFuse.Category UBound = " + CStr(UBound(MONFuse.Category))
    If (gB_findCMP_flag) Then DebugPrintLog "7... CMPFuse.Category UBound = " + CStr(UBound(CMPFuse.Category))
    If (gB_findUDRE_flag) Then DebugPrintLog "7... UDRE_Fuse.Category UBound = " + CStr(UBound(UDRE_Fuse.Category))
    If (gB_findUDRP_flag) Then DebugPrintLog "7... UDRP_Fuse.Category UBound = " + CStr(UBound(UDRP_Fuse.Category))
    If (gB_findCMPE_flag) Then DebugPrintLog "7... CMPE_Fuse.Category UBound = " + CStr(UBound(CMPE_Fuse.Category))
    If (gB_findCMPP_flag) Then DebugPrintLog "7... CMPP_Fuse.Category UBound = " + CStr(UBound(CMPP_Fuse.Category))
    
    ''''201811XX Update
    If (gB_findECID_flag) Then
        ''''build up the global Dictionary to speed up the Index search
        For i = 0 To UBound(ECIDFuse.Category)
            m_keyname = ECIDFuse.Category(i).Name
            Call eFuse_AddStoredIndex(eFuse_ECID, m_keyname, i)
        Next i
        '''--------------------------------------------------------------------------
        ReDim Preserve ECIDFuse.Category(CateCnt_ECID) ''''<MUST> final dimension
        '''--------------------------------------------------------------------------
        ''''Need to add one more category of "ECID_DEID" for the new table format
        n = CateCnt_ECID
        ECIDFuse.Category(n).index = n
        ECIDFuse.Category(n).Name = "ECID_DEID"
        ECIDFuse.Category(n).SeqStart = 0
        ECIDFuse.Category(n).SeqEnd = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).LSBbit
        ECIDFuse.Category(n).MSBbit = 0
        ECIDFuse.Category(n).LSBbit = ECIDFuse.Category(n).SeqEnd                                       ''''was=52
        ECIDFuse.Category(n).BitWidth = (ECIDFuse.Category(n).LSBbit - ECIDFuse.Category(n).MSBbit) + 1 ''''was=53
        ECIDFuse.Category(n).Stage = "CP1"
        ECIDFuse.Category(n).LoLMT = 7
        ECIDFuse.Category(n).HiLMT = ECIDFuse.Category(n).BitWidth ''''was=53
        ECIDFuse.Category(n).Resoultion = 0#
        ECIDFuse.Category(n).algorithm = "DEID"
        ECIDFuse.Category(n).comment = "ECID First DEID"
        ECIDFuse.Category(n).DefaultValue = 0
        ECIDFuse.Category(n).MSBFirst = "Y"
        CateCnt_ECID = CateCnt_ECID + 1
        '''--------------------------------------------------------------------------
        DebugPrintLog "11...ECIDFuse.Category Array Size = " + CStr(CateCnt_ECID)
        ''''ReDim Preserve ECIDFuse.Category(CateCnt_ECID - 1) ''''final dimension
        ''''add one more to the Index Dictionary
        m_keyname = ECIDFuse.Category(n).Name
        Call eFuse_AddStoredIndex(eFuse_ECID, m_keyname, n)
        '''--------------------------------------------------------------------------
    End If

    ''''201811XX
    If (gB_findCFG_flag) Then
        ''''build up the global Dictionary to speed up the Index search
        For i = 0 To UBound(CFGFuse.Category)
            m_keyname = CFGFuse.Category(i).Name
            Call eFuse_AddStoredIndex(eFuse_CFG, m_keyname, i)
        Next i
    End If
    
    If (gB_findUID_flag) Then
    End If
    If (gB_findUDR_flag) Then
        ''''build up the global Dictionary to speed up the Index search
        For i = 0 To UBound(UDRFuse.Category)
            m_keyname = UDRFuse.Category(i).Name
            Call eFuse_AddStoredIndex(eFuse_UDR, m_keyname, i)
        Next i
    End If
    If (gB_findSEN_flag) Then
    End If
    If (gB_findMON_flag) Then
        ''''build up the global Dictionary to speed up the Index search
        For i = 0 To UBound(MONFuse.Category)
            m_keyname = MONFuse.Category(i).Name
            Call eFuse_AddStoredIndex(eFuse_MON, m_keyname, i)
        Next i
    End If
    If (gB_findCMP_flag) Then
    End If
    If (gB_findUDRE_flag) Then
        For i = 0 To UBound(UDRE_Fuse.Category)
            m_keyname = UDRE_Fuse.Category(i).Name
            Call eFuse_AddStoredIndex(eFuse_UDRE, m_keyname, i)
        Next i
    End If
    If (gB_findUDRP_flag) Then
        For i = 0 To UBound(UDRP_Fuse.Category)
            m_keyname = UDRP_Fuse.Category(i).Name
            Call eFuse_AddStoredIndex(eFuse_UDRP, m_keyname, i)
        Next i
    End If
    If (gB_findCMPE_flag) Then
    End If
    If (gB_findCMPP_flag) Then
    End If

    ''''201811XX
    If (False) Then ''''debug purpose
        For j = 0 To UBound(CFGFuse.Category)
            m_keyname = CFGFuse.Category(j).Name
            k = eFuse_GetStoredIndex(eFuse_CFG, m_keyname)
            TheExec.Datalog.WriteComment m_keyname & " Index = " & k
        Next j
        
        For j = 0 To UBound(ECIDFuse.Category)
            m_keyname = ECIDFuse.Category(j).Name
            k = eFuse_GetStoredIndex(eFuse_ECID, m_keyname)
            TheExec.Datalog.WriteComment m_keyname & " Index = " & k
        Next j
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function auto_eFuseCategoryResult_Initialize() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuseCategoryResult_Initialize"
    
    Dim m_Site As Variant
    Dim i, k As Long
    Dim m_bitwidth As Long
    Dim m_startIdx As Long
    Dim m_tmpArr() As Long

    Dim m_resetFuseParam As EFuseCategoryParamResultSyntax
    
    ''''By this way to let all Fuse parameters to Nothing/CLear
    ''''Could choose any one of the members as the representative (.Decimal, .HexStr, ...)
    Set m_resetFuseParam.Decimal = Nothing
    
    If (gB_findECID_flag = True) Then
        For i = 0 To UBound(ECIDFuse.Category)
            With ECIDFuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "Y"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx - k ''''<Special> ECID
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

    If (gB_findCFG_flag = True) Then
        For i = 0 To UBound(CFGFuse.Category)
            With CFGFuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "N"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx + k
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

    If (gB_findUID_flag = True) Then
        For i = 0 To UBound(UIDFuse.Category)
            With UIDFuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "N"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx + k
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

    If (gB_findUDR_flag = True) Then
        For i = 0 To UBound(UDRFuse.Category)
            With UDRFuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "N"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx + k
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

    If (gB_findCMP_flag = True) Then
        For i = 0 To UBound(CMPFuse.Category)
            With CMPFuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "N"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx + k
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

    If (gB_findSEN_flag = True) Then
        For i = 0 To UBound(SENFuse.Category)
            With SENFuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "N"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx + k
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

    If (gB_findMON_flag = True) Then
        For i = 0 To UBound(MONFuse.Category)
            With MONFuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "N"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx + k
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

    ''''20171103 update
    If (gB_findUDRE_flag = True) Then
        For i = 0 To UBound(UDRE_Fuse.Category)
            With UDRE_Fuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "N"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx + k
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

    If (gB_findUDRP_flag = True) Then
        For i = 0 To UBound(UDRP_Fuse.Category)
            With UDRP_Fuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "N"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx + k
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

    If (gB_findCMPE_flag = True) Then
        For i = 0 To UBound(CMPE_Fuse.Category)
            With CMPE_Fuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "N"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx + k
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

    If (gB_findCMPP_flag = True) Then
        For i = 0 To UBound(CMPP_Fuse.Category)
            With CMPP_Fuse.Category(i)
                m_startIdx = .LSBbit ''''<MUST> be aligned with .BitArrWave
                .MSBFirst = "N"
                m_bitwidth = .BitWidth
                .HiLMT_R = 0#
                .LoLMT_R = 0#
                ReDim .DefValBitArr(m_bitwidth - 1)
                ReDim m_tmpArr(m_bitwidth - 1)
                For k = 0 To m_bitwidth - 1
                    m_tmpArr(k) = m_startIdx + k
                Next k
                For Each m_Site In TheExec.sites.Existing
                    .PatTestPass_Flag = True
                    .BitIndexWave.CreateConstant 0, m_bitwidth, DspLong
                    .BitIndexWave.Data = m_tmpArr
                Next m_Site
                .Read = m_resetFuseParam
                .Write = m_resetFuseParam
            End With
        Next i
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function ECIDIndex(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "ECIDIndex"

    Dim i As Long
    Dim match_Flag As Boolean
    match_Flag = False

    If (True) Then
        Dim m_keyname As String
        m_keyname = myStr
        ECIDIndex = eFuse_GetStoredIndex(eFuse_ECID, m_keyname)
        If (ECIDIndex >= 0) Then match_Flag = True
    Else
        For i = 0 To UBound(ECIDFuse.Category)
            If (UCase(myStr) = UCase(ECIDFuse.Category(i).Name)) Then
                ECIDIndex = i
                match_Flag = True
                Exit For
            End If
        Next i
    End If

    If (match_Flag = False) Then
        ECIDIndex = -1
        PrintDataLog "ECIDIndex:: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function CFGIndex(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CFGIndex"

    Dim i As Long
    Dim match_Flag As Boolean
    match_Flag = False
    
    If (True) Then
        Dim m_keyname As String
        m_keyname = myStr
        CFGIndex = eFuse_GetStoredIndex(eFuse_CFG, m_keyname)
        If (CFGIndex >= 0) Then match_Flag = True
    Else
        For i = 0 To UBound(CFGFuse.Category)
            If (UCase(myStr) = UCase(CFGFuse.Category(i).Name)) Then
                CFGIndex = i
                match_Flag = True
                Exit For
            End If
        Next i
    End If
    
    If (match_Flag = False) Then
        CFGIndex = -1
         If Not (myStr Like "ids_*_85") Then
        PrintDataLog "CFGIndex:: <" + myStr + ">, it's NOT existed in the Category."
        End If
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function UIDIndex(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "UIDIndex"

    Dim i As Long
    Dim match_Flag As Boolean

    match_Flag = False
    For i = 0 To UBound(UIDFuse.Category)
        If (UCase(myStr) = UCase(UIDFuse.Category(i).Name)) Then
            UIDIndex = i
            match_Flag = True
            Exit For
        End If
    Next i

    If (match_Flag = False) Then
        UIDIndex = -1
        PrintDataLog "UIDIndex:: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function UDRIndex(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "UDRIndex"

    Dim i As Long
    Dim match_Flag As Boolean
    match_Flag = False
    
    If (True) Then
        Dim m_keyname As String

        m_keyname = myStr
        UDRIndex = eFuse_GetStoredIndex(eFuse_UDR, m_keyname)
        If (UDRIndex >= 0) Then match_Flag = True
    Else
        For i = 0 To UBound(UDRFuse.Category)
            If (UCase(myStr) = UCase(UDRFuse.Category(i).Name)) Then
                UDRIndex = i
                match_Flag = True
                Exit For
            End If
        Next i
    End If

    If (match_Flag = False) Then
        UDRIndex = -1
        PrintDataLog "UDRIndex:: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function SENIndex(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "SENIndex"

    Dim i As Long
    Dim match_Flag As Boolean

    match_Flag = False
    For i = 0 To UBound(SENFuse.Category)
        If (UCase(myStr) = UCase(SENFuse.Category(i).Name)) Then
            SENIndex = i
            match_Flag = True
            Exit For
        End If
    Next i

    If (match_Flag = False) Then
        SENIndex = -1
        PrintDataLog "SENIndex:: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function MONIndex(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "MONIndex"

    Dim i As Long
    Dim match_Flag As Boolean
    match_Flag = False
    
    If (True) Then
        Dim m_keyname As String
        m_keyname = myStr
        MONIndex = eFuse_GetStoredIndex(eFuse_MON, m_keyname)
        If (MONIndex >= 0) Then match_Flag = True
    Else
        For i = 0 To UBound(MONFuse.Category)
            If (UCase(myStr) = UCase(MONFuse.Category(i).Name)) Then
                MONIndex = i
                match_Flag = True
                Exit For
            End If
        Next i
    End If

    If (match_Flag = False) Then
        MONIndex = -1
        PrintDataLog "MONIndex:: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function CMPIndex(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CMPIndex"

    Dim i As Long
    Dim match_Flag As Boolean

    match_Flag = False
'    For i = 0 To UBound(CMPFuse.Category)
'        If (UCase(myStr) = UCase(CMPFuse.Category(i).Name)) Then
'            CMPIndex = i
'            match_Flag = True
'            Exit For
'        End If
'    Next i

    If (True) Then
        Dim m_keyname As String
        m_keyname = myStr
        CMPIndex = eFuse_GetStoredIndex(eFuse_CMP, m_keyname)
        If (CMPIndex >= 0) Then match_Flag = True
    Else
        For i = 0 To UBound(CMPFuse.Category)
            If (UCase(myStr) = UCase(CMPFuse.Category(i).Name)) Then
                CMPIndex = i
                match_Flag = True
                Exit For
            End If
        Next i
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function CFGTabIndex(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CFGTabIndex"

    Dim i As Long
    Dim match_Flag As Boolean
    match_Flag = False
    
    If (True) Then
        Dim m_keyname As String
        m_keyname = myStr
        CFGTabIndex = eFuse_GetStoredIndex(eFuse_CFGTab, m_keyname)
        If (CFGTabIndex >= 0) Then match_Flag = True
    Else
        For i = 0 To UBound(CFGTable.Category)
            If (UCase(myStr) = UCase(CFGTable.Category(i).pkgName)) Then
                CFGTabIndex = i
                match_Flag = True
                Exit For
            End If
        Next i
    End If

    If (match_Flag = False) Then
        CFGTabIndex = -1
        PrintDataLog "CFGTabIndex:: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''Here bitstr (default) is [MSB......LSB]
''''20150915, update with the optional bitstrM_flag to True or False
Public Function auto_bitStr2Dec(BitStr As String, Optional bitstrM_flag As Boolean = True) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_bitStr2Dec"
    
    ''''EX:
    ''''bitstr=11001, m_dec=25
    
    Dim i As Long
    Dim m_dec As Long
    Dim BitWidth As Long
    
    BitWidth = Len(BitStr)
    m_dec = 0
    
    ''''case: bitstr is [LSB...MSB]
    ''''Then set bitstrM_flag to Fasle, bitstr should be reversed to [MSB...LSB]
    If (bitstrM_flag = False) Then
        BitStr = StrReverse(BitStr)
    End If
    
    
    ''''<NOTICE>
    ''''if BitWidth >31, it will result in an overflow error message and supposedy it's a reserved bits.
    If (BitWidth <= 31) Then
        For i = 0 To BitWidth - 1
            m_dec = m_dec + CLng(Mid(BitStr, i + 1, 1)) * (2 ^ (BitWidth - 1 - i))
        Next i
    Else
        ''supposedy it's a reserved bits and all '0'.
        m_dec = 0
    End If

    auto_bitStr2Dec = m_dec
    
    ''Debug.Print "bitstr=" + bitstr + ", Dec=" + CStr(m_dec)

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
      
End Function

'''' use to prevent the hang up issue if the iEDA register key has the wrong format.
Public Function auto_checkIEDAString(inStr1 As String) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_checkIEDAString"
    
    Dim k As Long
    Dim tmpA() As String
    Dim tmpStr As String
    Dim tmpA_size As Long

    tmpA = Split(inStr1, ",")
    tmpA_size = UBound(tmpA) + 1
    
    ''Debug.Print "tmpA size =" & tmpA_size
    tmpStr = ""
    If (tmpA_size < TheExec.sites.Existing.Count) Then
        If (tmpA_size = 0) Then
            tmpStr = "NA"
            For k = 1 To (TheExec.sites.Existing.Count - tmpA_size - 1)
                tmpStr = tmpStr + ",NA"
            Next k
        Else
            For k = 1 To (TheExec.sites.Existing.Count - tmpA_size)
                tmpStr = tmpStr + ",NA"
            Next k
        End If
        TheExec.Datalog.WriteComment funcName + ":: original = " + inStr1
        inStr1 = inStr1 + tmpStr
        TheExec.Datalog.WriteComment Space(23) + "  update = " + inStr1 + " ......could have the site sequence problem (case1) !!!" + vbCrLf
    
    ElseIf (tmpA_size > TheExec.sites.Existing.Count) Then ''''should not have this case
        tmpStr = ""
        For k = 0 To TheExec.sites.Existing.Count - 1
            If (k = (TheExec.sites.Existing.Count - 1)) Then
                tmpStr = tmpStr + tmpA(k)
            Else
                tmpStr = tmpStr + tmpA(k) + ","
            End If
        Next k
        TheExec.Datalog.WriteComment funcName + ":: original = " + inStr1
        inStr1 = tmpStr
        TheExec.Datalog.WriteComment Space(23) + "  update = " + inStr1 + " ......could have the site sequence problem (case2) !!!" + vbCrLf

    End If
    
    auto_checkIEDAString = inStr1
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20151119 update for New bincut method2
Public Function auto_Vbin_to_VfuseStr_New(vbin As Variant, BitWidth As Long, VfuseStr As String, m_stepVoltage As Double) As Variant

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Vbin_to_VfuseStr_New"
    
    '***** Step 1 :  Translate Vbin to Vfuse by equation on test plan  *******
    'Translation from Vbin to Vfuse by the alogrithm written in test plan
    'Bin Voltage =(Base Fuse +1)*25 + (Bin Fuse * 5)
    'In other word, Bin Fuse = (Bin Voltage -(Base Fuse +1)*25)/5

    Dim FuseCode As Long
    Dim FuseStr As Variant
    Dim m_binarr() As Long
    Dim i As Long
    Dim m_upperValue As Long

    m_upperValue = 2 ^ BitWidth - 1
    VfuseStr = ""

    If vbin > 0 Then
        ''''gD_VBaseFuse = gD_BaseVoltage / gD_BaseStepVoltage - 1
        ''''so gD_BaseVoltage = (gD_VBaseFuse + 1) * gD_BaseStepVoltage
        FuseCode = CeilingValue((vbin - gD_BaseVoltage) / m_stepVoltage, 1) ''''20160608 update

        If (FuseCode > m_upperValue) Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: FuseCode(" & FuseCode & ") > " & m_upperValue & " (upperValue)"
            GoTo errHandler
        End If
        '***** Step 2 :  Translate decimal to binary string *******
        VfuseStr = auto_Dec2Bin_EFuse(FuseCode, BitWidth, m_binarr) ''''was auto_Dec2BinArr(), 20160608 update
    Else
        For i = 1 To BitWidth
            VfuseStr = VfuseStr + "0"
        Next i
        FuseCode = 0
        If vbin < 0 Then
            TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: input value <0 (" & vbin & ") and force it to zero."
        End If
    End If
    auto_Vbin_to_VfuseStr_New = FuseCode
    'Debug.Print funcName + ":: " + VfuseStr + "  (FuseCode=" + CStr(VfuseStr) + ", Vbin=" + CStr(Vbin) + ")"

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150612, update
''''20171117 update to gate the negative input (need to be verified)
Public Function auto_Dec2Bin_EFuse(ByVal n As Long, BitWidth As Long, ByRef BinArray() As Long) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Dec2Bin_EFuse"

    ''Debug.Print "n = " & n & ", bitwidth=" & BitWidth
''''-----------------------------------
''''<Example>
''''n = 11, bitwidth=6
''''BinArray [0] = 1
''''BinArray [1] = 1
''''BinArray [2] = 0
''''BinArray [3] = 1
''''BinArray [4] = 0
''''BinArray [5] = 0
''''m_bitstrM [MSB...LSB] = 001011
''''-----------------------------------

    Dim i As Long
    Dim m_bitStrM As String
    
    ''Initialize the content of array
    ReDim BinArray(BitWidth - 1) ''''BinArray[0] is LSB
    m_bitStrM = ""

''''    ''''20171117 update to gate the negative input
''''    If (n < 0) Then
''''        TheExec.AddOutput "<Error> " + funcName + ":: the input n=" + CStr(n) + " is the negative value."
''''        TheExec.Datalog.WriteComment "<Error> " + funcName + ":: the input n=" + CStr(n) + " is the negative value."
''''        GoTo errHandler
''''    End If

    For i = 0 To BitWidth - 1
        BinArray(i) = 0
        If (n Mod 2) Then
            BinArray(i) = 1
        Else
            BinArray(i) = 0
        End If
        m_bitStrM = CStr(BinArray(i)) + m_bitStrM ''''[MSB...LSB]
        n = Fix(n / 2)
        ''Debug.Print "BinArray[" & i & "] = " & BinArray(i)
    Next i
    ''Debug.Print "m_bitstrM[MSB...LSB] = " & m_bitstrM
    
    auto_Dec2Bin_EFuse = m_bitStrM

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150825 New Function to Replace auto_(CFG/UID/UDR/SEN/MON/CMP)_Bin2DecStr()
''''20160503 Update for the data which exceeds over 32bits by using Double
''''20160927 Update, 20171103 update
''''20180522 Update to handle over 1023 bits
Public Function auto_eFuse_Bin2DecStr(ByVal FuseType As String, idx As Long, SrcArr() As Long, LSBbit As Long, MSBbit As Long) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Bin2DecStr"
    
    Dim m_HexStr As String
    Dim m_bitStrM As String
    Dim m_decimal As Long
    Dim i As Long, j As Long
    Dim m_bitsum As Long
    Dim ss As Variant
    Dim BitWidth As Long
    Dim m_dbl As Variant ''''20160503, 20180522 Double->Variant
    Dim m_tmpStr As String
    Dim m_dbl_overrange As Boolean

    FuseType = UCase(Trim(FuseType))
    ss = TheExec.sites.SiteNumber
    BitWidth = Abs(MSBbit - LSBbit) + 1
    
    m_HexStr = ""
    m_bitStrM = ""
    m_decimal = 0
    m_bitsum = 0
    m_dbl = 0
    m_dbl_overrange = False
    m_tmpStr = "<Error> " + funcName + ":: BitWidth is over 1023 and not Zero (" + FuseType + ", idx=" + CStr(idx) + ")"
    ''''--------------------------------------------------------------------------------------------
    ''''<Notice> It should only handle up to 31bits
    ''''CLng(2^31) will be Overflow (Run Time Error)
    ''''CDbl(2^1024) will be Overflow (Run Time Error)
    ''''--------------------------------------------------------------------------------------------
    If (LSBbit <= MSBbit) Then
        For i = LSBbit To MSBbit
            m_bitStrM = CStr(SrcArr(i)) + m_bitStrM ''''translate to [MSB(end)......LSB(start)]
            m_bitsum = m_bitsum + SrcArr(i)
            j = Abs(i - LSBbit)
            
            ''''20180522 update for bits over 1023
            If (j < 1024) Then
                m_dbl = m_dbl + SrcArr(i) * CDbl(2 ^ j) ''''<MUST> using CDbl(2 ^ j) to avoid overflow, up to 1024bits(j<1024)
            ElseIf (j >= 1024 And SrcArr(i) = 0) Then
                m_dbl = m_dbl + 0
            Else
                m_dbl_overrange = True
                ''TheExec.AddOutput m_tmpStr
                ''TheExec.Datalog.WriteComment m_tmpStr
                ''GoTo errHandler
            End If
        Next i
    Else
        ''''case:: LSBbit > MSBbit
        For i = LSBbit To MSBbit Step -1
            m_bitStrM = CStr(SrcArr(i)) + m_bitStrM ''''translate to [MSB(end)......LSB(start)]
            m_bitsum = m_bitsum + SrcArr(i)
            j = Abs(i - LSBbit)
                                 
            ''''20180522 update for bits over 1023
            If (j < 1024) Then
                m_dbl = m_dbl + SrcArr(i) * CDbl(2 ^ j) ''''<MUST> using CDbl(2 ^ j) to avoid overflow, up to 1024bits(j<1024)
            ElseIf (j >= 1024 And SrcArr(i) = 0) Then
                m_dbl = m_dbl + 0
            Else
                m_dbl_overrange = True
                ''TheExec.AddOutput m_tmpStr
                ''TheExec.Datalog.WriteComment m_tmpStr
                ''GoTo errHandler
            End If
        Next i
    End If
    m_HexStr = auto_BinStr2HexStr(m_bitStrM, 1) ''''20170911 update
    If (m_dbl_overrange = True) Then m_dbl = "0x" + m_HexStr ''''20180522 update
    ''''--------------------------------------------------------------------------------------------

    If (FuseType = "ECID") Then
        ''''20170911 update and support
        ECIDFuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        ECIDFuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        ECIDFuse.Category(idx).Read.Decimal(ss) = m_dbl
        ECIDFuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        ECIDFuse.Category(idx).Read.Value(ss) = m_dbl
        ECIDFuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        ECIDFuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr

    ElseIf (FuseType = "CFG") Then
        CFGFuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        CFGFuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        CFGFuse.Category(idx).Read.Decimal(ss) = m_dbl
        CFGFuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        CFGFuse.Category(idx).Read.Value(ss) = m_dbl
        CFGFuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        CFGFuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr ''''20160907 update

    ElseIf (FuseType = "UID") Then
        UIDFuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        UIDFuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        UIDFuse.Category(idx).Read.Decimal(ss) = m_dbl
        UIDFuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        UIDFuse.Category(idx).Read.Value(ss) = m_dbl
        UIDFuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        UIDFuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr ''''20160907 update

    ElseIf (FuseType = "UDR") Then
        UDRFuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        UDRFuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRFuse.Category(idx).Read.Decimal(ss) = m_dbl
        UDRFuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        UDRFuse.Category(idx).Read.Value(ss) = m_dbl
        UDRFuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        UDRFuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr ''''20160907 update

    ElseIf (FuseType = "SEN") Then
        SENFuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        SENFuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        SENFuse.Category(idx).Read.Decimal(ss) = m_dbl
        SENFuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        SENFuse.Category(idx).Read.Value(ss) = m_dbl
        SENFuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        SENFuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr ''''20160907 update

    ElseIf (FuseType = "MON") Then
        MONFuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        MONFuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        MONFuse.Category(idx).Read.Decimal(ss) = m_dbl
        MONFuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        MONFuse.Category(idx).Read.Value(ss) = m_dbl
        MONFuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        MONFuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr ''''20160907 update

    ElseIf (FuseType = "CMP") Then
        CMPFuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        CMPFuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        CMPFuse.Category(idx).Read.Decimal(ss) = m_dbl
        CMPFuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        CMPFuse.Category(idx).Read.Value(ss) = m_dbl
        CMPFuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        CMPFuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr ''''20160907 update

    ElseIf (FuseType = "UDRE") Then
        UDRE_Fuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        UDRE_Fuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRE_Fuse.Category(idx).Read.Decimal(ss) = m_dbl
        UDRE_Fuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        UDRE_Fuse.Category(idx).Read.Value(ss) = m_dbl
        UDRE_Fuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        UDRE_Fuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr ''''20160907 update

    ElseIf (FuseType = "UDRP") Then
        UDRP_Fuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        UDRP_Fuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRP_Fuse.Category(idx).Read.Decimal(ss) = m_dbl
        UDRP_Fuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        UDRP_Fuse.Category(idx).Read.Value(ss) = m_dbl
        UDRP_Fuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        UDRP_Fuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr ''''20160907 update

    ElseIf (FuseType = "CMPE") Then
        CMPE_Fuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        CMPE_Fuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        CMPE_Fuse.Category(idx).Read.Decimal(ss) = m_dbl
        CMPE_Fuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        CMPE_Fuse.Category(idx).Read.Value(ss) = m_dbl
        CMPE_Fuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        CMPE_Fuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr ''''20160907 update

    ElseIf (FuseType = "CMPP") Then
        CMPP_Fuse.Category(idx).Read.BitStrM(ss) = m_bitStrM
        CMPP_Fuse.Category(idx).Read.BitStrL(ss) = StrReverse(m_bitStrM)
        CMPP_Fuse.Category(idx).Read.Decimal(ss) = m_dbl
        CMPP_Fuse.Category(idx).Read.BitSummation(ss) = m_bitsum
        CMPP_Fuse.Category(idx).Read.Value(ss) = m_dbl
        CMPP_Fuse.Category(idx).Read.ValStr(ss) = CStr(m_dbl)
        CMPP_Fuse.Category(idx).Read.HexStr(ss) = "0x" + m_HexStr ''''20160907 update

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
    End If
    ''''--------------------------------------------------------------------------------------------

    If (BitWidth <= 31) Then
        auto_eFuse_Bin2DecStr = CStr(m_decimal)
    Else
        auto_eFuse_Bin2DecStr = CStr(m_dbl) ''20160927 update, was m_hexStr
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function TurnOnEfusePwrPins(powerPin As String, _
                                   Optional v As Double = 1.8, _
                                   Optional i_rng As Double = 0.2, _
                                   Optional wait_before_gate As Double = 0.001, _
                                   Optional wait_after_gate As Double = 0.002, _
                                   Optional Steps As Long = 10, _
                                   Optional RiseTime As Double = 0.001)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "TurnOnEfusePwrPins"

    '******************** 2015/1/16 Laba ********************
    'ECID Supply : VDD18_EFUSE0
    'HDCP Keys, UID Keys, Sensor Trim Values, Config : VDD18_EFUSE1
    'SOC BIRA1, SOC BIRA2, CPU UDR+BIRA, GFX BIRA : VDD18_EFUSE2
    
    Dim m_currVolt As Double
    
    ''''auto_eFuse_pwr_on_i_meter_DCVS(pin As String, v As Double, i_rng As Double, wait_before_gate As Double, wait_after_gate As Double, steps As Long, RiseTime As Double)
    ''''auto_eFuse_pwr_on_i_meter_DCVS PowerPin, vpwr, 0.2, 0.001, 0.002, 10, 0.001
    auto_eFuse_pwr_on_i_meter_DCVS powerPin, v, i_rng, wait_before_gate, wait_after_gate, Steps, RiseTime
    
    m_currVolt = TheHdw.DCVS.Pins(powerPin).Voltage.Main.Value
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment funcName + " :: " + UCase(powerPin) + " = " + Format(m_currVolt, "0.000")
    TheExec.Datalog.WriteComment ""

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function TurnOffEfusePwrPins(powerPin As String, _
                                   Optional v As Double = 1.8, _
                                   Optional i_rng As Double = 0.2, _
                                   Optional wait_before_gate As Double = 0.001, _
                                   Optional wait_after_gate As Double = 0.002, _
                                   Optional Steps As Long = 10, _
                                   Optional RiseTime As Double = 0.001)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "TurnOffEfusePwrPins"

    Dim m_currVolt As Double
    ''auto_eFuse_pwr_off_i_meter_DCVS CurrentVoltage, 1.8, 0.2, 0.001, 0.002, 10, 0.001
    
    auto_eFuse_pwr_off_i_meter_DCVS powerPin, v, i_rng, wait_before_gate, wait_after_gate, Steps, RiseTime
    
    m_currVolt = TheHdw.DCVS.Pins(powerPin).Voltage.Main.Value
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment funcName + " :: " + UCase(powerPin) + " = " + Format(m_currVolt, "0.000")
    TheExec.Datalog.WriteComment ""
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function



Public Function FormatNumeric(num As Variant, length As Long) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "FormatNumeric"
    
    ''''Example
    ''''----------------------------------------
    '''' length > 0  is to right shift
    '''' length < 0  is to left  shift
    ''''----------------------------------------
    ''''FormatNumeric(123456, 8) + "...end"
    ''''  123456...end
    ''''
    ''''FormatNumeric(123456,-8) + "...end"
    ''''123456  ...end
    ''''
    ''''----------------------------------------
    
    Dim numStr As String
    Dim tmpLen As Long
    Dim spcLen As Long
    
    numStr = CStr(num)
    tmpLen = Len(numStr)
    
    If (tmpLen > Abs(length)) Then
        spcLen = 0
    Else
        spcLen = Abs(length) - tmpLen
    End If
    
    If (length < 0) Then   ''''number shift to the very left
        FormatNumeric = numStr + Space(spcLen)
    Else ''''default: shift to the very right
        FormatNumeric = Space(spcLen) + numStr
    End If

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''2015-04-23
''''Here it's used to update the column length in the datalog.
Public Function UpdateDLogColumns(tsNameWidth As Variant) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "UpdateDLogColumns"

    ''''20170217 update
    If (gB_newDlog_Flag) Then Exit Function

    tsNameWidth = CLng(tsNameWidth)
    TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
    With TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric
        .testName.Enable = True
        .testName.Width = tsNameWidth
        .Pin.Enable = True
        .Pin.Width = 25
    End With
    With TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional
        .Pattern.Enable = True
        .Pattern.Width = 128 '.Pattern.DefaultWidth
        .testName.Enable = True
        .testName.Width = tsNameWidth
    End With
    TheExec.Datalog.ApplySetup

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''2015-07-02
''''Here it's used to disable the column length in the datalog.
Public Function UpdateDLogColumns__False()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "UpdateDLogColumns__False"

    ''''20170217 update
    If (gB_newDlog_Flag) Then Exit Function

    TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = False
    TheExec.Datalog.ApplySetup

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20150625 New Function
''''20160201 update code, add optional "chkStage"
''''20160613 update As Long to As Variant for the function()
''''20160624 update to have the optional of "decimal" for "default or real"
''''20171211 update to have the capabilty to check if it's binary string input.
Public Function auto_eFuse_Get_DefaultRealDecimal(ByVal FuseType As String, m_catename As String, defreal As String, defval As Variant, _
                    Optional chkStage As Boolean = True) As Variant

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Get_DefaultRealDecimal"

    Dim m_decimal As Variant ''''20160613 update
    Dim m_defval As Variant
    
    defreal = LCase(defreal)
    
    ''''20171211, Here it's used to judge default value if is Hex or Binary
    ''m_defval = auto_checkDefaultValue(defval)
    ''''The above is masked out because the DefaultValue has been processed in the routine auto_[fuse]Constant_Initialize().
    ''''so we do not need to do it again (it's redundant).
    m_defval = defval
    
    If (defreal = "default" Or defreal = "decimal") Then ''''20160624 update
        m_decimal = m_defval
    ElseIf (defreal = "real") Then
        m_decimal = auto_eFuse_GetWriteDecimal(FuseType, m_catename, False)
    ElseIf (defreal Like "safe*voltage") Then
        m_decimal = m_defval
    Else
        ''''if unknown state
        TheExec.Datalog.WriteComment "<WARNING> auto_eFuse_Get_DefaultRealDecimal::" + FuseType + "Fuse set Decimal=0 in this unKnown state, should be one of (Default/Real/Decimal/Safe Voltage)"
        m_decimal = 0
    End If
    
    ''''20150707 update, 20160201 update
    If (chkStage = True) Then
        If (auto_eFuse_chkStage(FuseType, m_catename) = 0) Then
            m_decimal = 0
        End If
    End If

    If (m_decimal < 0) Then
        m_decimal = 0  'prevent negative value
        ''''201807XX update to meet old/new method
        Call auto_eFuse_SetWriteDecimal(FuseType, m_catename, m_decimal, False, chkStage)
    End If

    auto_eFuse_Get_DefaultRealDecimal = m_decimal

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150625 New Function
''''20160201 update code with optional 'chkStage'
''''20160616 update as Variant
''''20180731 update with new method, skip the chkStage
Public Function auto_eFuse_SetWriteDecimal(ByVal FuseType As String, m_catename As String, ByVal m_value As Variant, _
                                           Optional showPrint As Boolean = True, _
                                           Optional chkStage As Boolean = False) As Variant

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_SetWriteDecimal"

    Dim m_len As Long
    Dim m_decimal As Variant
    Dim m_dlogstr As String
    Dim ss As Variant

    Dim m_tmpStr As String
    Dim m_BinStr As String
    Dim m_HexStr As String
    Dim m_idx As Long
    Dim m_bitStrM As String
    Dim m_binarr() As Long
    Dim m_bitwidth As Long
    Dim m_FuseWrite As EFuseCategoryParamResultSyntax

    ss = TheExec.sites.SiteNumber
    FuseType = UCase(Trim(FuseType))

    If (FuseType = "ECID") Then
        m_idx = ECIDIndex(m_catename)
        m_bitwidth = ECIDFuse.Category(m_idx).BitWidth
        m_FuseWrite = ECIDFuse.Category(m_idx).Write
    ElseIf (FuseType = "CFG") Then
        m_idx = CFGIndex(m_catename)
        m_bitwidth = CFGFuse.Category(m_idx).BitWidth
        m_FuseWrite = CFGFuse.Category(m_idx).Write
    ElseIf (FuseType = "UID") Then
        m_idx = UIDIndex(m_catename)
        m_bitwidth = UIDFuse.Category(m_idx).BitWidth
        m_FuseWrite = UIDFuse.Category(m_idx).Write
    ElseIf (FuseType = "UDR") Then
        m_idx = UDRIndex(m_catename)
        m_bitwidth = UDRFuse.Category(m_idx).BitWidth
        m_FuseWrite = UDRFuse.Category(m_idx).Write
    ElseIf (FuseType = "SEN") Then
        m_idx = SENIndex(m_catename)
        m_bitwidth = SENFuse.Category(m_idx).BitWidth
        m_FuseWrite = SENFuse.Category(m_idx).Write
    ElseIf (FuseType = "MON") Then
        m_idx = MONIndex(m_catename)
        m_bitwidth = MONFuse.Category(m_idx).BitWidth
        m_FuseWrite = MONFuse.Category(m_idx).Write
    ElseIf (FuseType = "CMP") Then
        m_idx = CMPIndex(m_catename)
        m_bitwidth = CMPFuse.Category(m_idx).BitWidth
        m_FuseWrite = CMPFuse.Category(m_idx).Write
    ElseIf (FuseType = "UDRE") Then
        m_idx = UDRE_Index(m_catename)
        m_bitwidth = UDRE_Fuse.Category(m_idx).BitWidth
        m_FuseWrite = UDRE_Fuse.Category(m_idx).Write
    ElseIf (FuseType = "UDRP") Then
        m_idx = UDRP_Index(m_catename)
        m_bitwidth = UDRP_Fuse.Category(m_idx).BitWidth
        m_FuseWrite = UDRP_Fuse.Category(m_idx).Write
    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP)"
        GoTo errHandler
        ''''nothing
    End If

    ''''20160620 update, if it's Hex, it MUST be with the prefix "0x" or "x"
    If (auto_isHexString(CStr(m_value))) Then
        'm_hexStr = m_value
        If (auto_chkHexStr_isOver7FFFFFFF(CStr(m_value)) = False) Then
            m_value = Replace(UCase(CStr(m_value)), "0X", "", 1, 1)
            m_value = Replace(UCase(CStr(m_value)), "X", "", 1, 1) ''''20171211 update
            m_value = Replace(UCase(CStr(m_value)), "_", "")       ''''20171211 update
            m_HexStr = "0x" + CStr(m_value)
            m_value = CLng("&H" & m_value) ''''Here it's Hex2Dec
        Else
            ''''<MUST> keep prefix "0x" or "x"
            m_value = Replace(UCase(CStr(m_value)), "0X", "", 1, 1)
            m_value = Replace(UCase(CStr(m_value)), "X", "", 1, 1) ''''20171211 update
            m_value = Replace(UCase(CStr(m_value)), "_", "")       ''''20171211 update
            m_HexStr = "0x" + CStr(m_value)
        End If
        m_decimal = m_value

    ElseIf (auto_isBinaryString(CStr(m_value))) Then ''''20171211 add
        m_BinStr = Replace(UCase(CStr(m_value)), "B", "", 1, 1) ''''remove the first "B" character
        m_BinStr = Replace(m_BinStr, "_", "")                   ''''for case like as b0001_0101
        
        ''''convert to Hex with the prefix "0x"
        m_HexStr = "0x" + auto_BinStr2HexStr(m_BinStr, 1)
        If (auto_chkHexStr_isOver7FFFFFFF(CStr(m_HexStr)) = False) Then
            m_HexStr = Replace(UCase(CStr(m_HexStr)), "0X", "", 1, 1)
            m_value = CLng("&H" & m_HexStr) ''''Here it's Hex2Dec
            m_decimal = m_value
        Else
            ''''<MUST> keep prefix "0x" or "x"
            m_decimal = m_HexStr
        End If

    Else
        ''''201808XX update to avoid "bincut" case
        m_HexStr = auto_Value2HexStr(m_value, m_bitwidth)
        If (m_HexStr Like "0x0*" And IsNumeric(m_value) = False) Then
            ''''case m_value = "bincut" Or "na" Or "n/a"
            m_tmpStr = Replace(UCase(CStr(m_HexStr)), "0X", "", 1, 1)
            m_decimal = CDbl("&H" & m_tmpStr) ''''Here it's Hex2Dec
        Else
            m_decimal = m_value
        End If

    End If
  
    ''''20180711 New for DSPWave
    m_bitStrM = auto_HexStr2BinStr_EFUSE(m_HexStr, m_bitwidth, m_binarr)
    
'    ''''20150707 update
'    If (chkStage = True) Then
'        If (auto_eFuse_chkStage(fusetype, m_catename) = 0) Then
'            m_decimal = 0
'        End If
'    End If

    Dim m_tmpWave As New DSPWave
    m_tmpWave.CreateConstant 0, m_bitwidth, DspLong
    m_tmpWave(ss).Data = m_binarr

    With m_FuseWrite
        .Decimal = m_decimal
        .BitArrWave = m_tmpWave.Copy
        .HexStr = m_HexStr
        .Value = m_decimal
        .BitSummation = m_tmpWave.CalcSum
        .BitStrM = m_bitStrM
        .BitStrL = StrReverse(m_bitStrM)
        .ValStr = CStr(m_decimal)
    End With
        
    If (FuseType = "ECID") Then
        ECIDFuse.Category(m_idx).Write = m_FuseWrite

    ElseIf (FuseType = "CFG") Then
        CFGFuse.Category(m_idx).Write = m_FuseWrite

    ElseIf (FuseType = "UID") Then
        UIDFuse.Category(m_idx).Write = m_FuseWrite

    ElseIf (FuseType = "UDR") Then
        UDRFuse.Category(m_idx).Write = m_FuseWrite

    ElseIf (FuseType = "SEN") Then
        SENFuse.Category(m_idx).Write = m_FuseWrite

    ElseIf (FuseType = "MON") Then
        MONFuse.Category(m_idx).Write = m_FuseWrite

    ElseIf (FuseType = "CMP") Then
        CMPFuse.Category(m_idx).Write = m_FuseWrite

    ElseIf (FuseType = "UDRE") Then
        UDRE_Fuse.Category(m_idx).Write = m_FuseWrite

    ElseIf (FuseType = "UDRP") Then
        UDRP_Fuse.Category(m_idx).Write = m_FuseWrite

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP)"
        GoTo errHandler
        ''''nothing
    End If

    auto_eFuse_SetWriteDecimal = m_decimal

    ''''20171211 update
    If (m_decimal < 0) Then
        GoTo errHandler
    End If

    If (showPrint) Then
        m_len = auto_eFuse_GetCatenameMaxLen(FuseType)
        FuseType = FormatNumeric(FuseType, 4)
        m_dlogstr = vbTab & "Site(" + CStr(ss) + ") " + FuseType + FormatNumeric("Fuse SetWriteDecimal", -25)
        m_dlogstr = m_dlogstr + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_decimal, -20)
        m_dlogstr = m_dlogstr + FormatNumeric(" [" + m_HexStr + "]", -1)
        TheExec.Datalog.WriteComment m_dlogstr
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150625 New Function
Public Function auto_eFuse_GetWriteDecimal(ByVal FuseType As String, m_catename As String, Optional showPrint As Boolean = True) As Variant
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_GetWriteDecimal"
    
    Dim m_len As Long
    Dim m_decimal As Variant
    Dim m_dlogstr As String
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber

    FuseType = UCase(Trim(FuseType))
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)

    If (FuseType = "ECID") Then
        m_decimal = ECIDFuse.Category(ECIDIndex(m_catename)).Write.Decimal(ss)

    ElseIf (FuseType = "CFG") Then
        m_decimal = CFGFuse.Category(CFGIndex(m_catename)).Write.Decimal(ss)

    ElseIf (FuseType = "UID") Then
        m_decimal = UIDFuse.Category(UIDIndex(m_catename)).Write.Decimal(ss)

    ElseIf (FuseType = "UDR") Then
        m_decimal = UDRFuse.Category(UDRIndex(m_catename)).Write.Decimal(ss)

    ElseIf (FuseType = "SEN") Then
        m_decimal = SENFuse.Category(SENIndex(m_catename)).Write.Decimal(ss)
        
    ElseIf (FuseType = "MON") Then
        m_decimal = MONFuse.Category(MONIndex(m_catename)).Write.Decimal(ss)

    ElseIf (FuseType = "CMP") Then
        m_decimal = CMPFuse.Category(CMPIndex(m_catename)).Write.Decimal(ss)

    ElseIf (FuseType = "UDRE") Then
        m_decimal = UDRE_Fuse.Category(UDRE_Index(m_catename)).Write.Decimal(ss)

    ElseIf (FuseType = "UDRP") Then
        m_decimal = UDRP_Fuse.Category(UDRP_Index(m_catename)).Write.Decimal(ss)

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP)"
        GoTo errHandler
        ''''nothing
    End If

    auto_eFuse_GetWriteDecimal = m_decimal

    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + FormatNumeric("Fuse GetWriteDecimal", -25)
        m_dlogstr = vbTab & "Site(" + CStr(ss) + ") " + FuseType + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_decimal, -10)
        TheExec.Datalog.WriteComment m_dlogstr
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150625 New Function
Public Function auto_eFuse_SetReadDecimal(ByVal FuseType As String, m_catename As String, m_value As Variant, Optional showPrint As Boolean = True) As Long
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_SetReadDecimal"

    Dim m_len As Long
    Dim m_decimal As Long
    Dim m_dlogstr As String
    Dim ss As Variant

    ss = TheExec.sites.SiteNumber
    FuseType = UCase(Trim(FuseType))
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)
    m_decimal = CLng(m_value)

    If (FuseType = "ECID") Then
        ECIDFuse.Category(ECIDIndex(m_catename)).Read.Decimal(ss) = m_decimal

    ElseIf (FuseType = "CFG") Then
        CFGFuse.Category(CFGIndex(m_catename)).Read.Decimal(ss) = m_decimal

    ElseIf (FuseType = "UID") Then
        UIDFuse.Category(UIDIndex(m_catename)).Read.Decimal(ss) = m_decimal

    ElseIf (FuseType = "UDR") Then
        UDRFuse.Category(UDRIndex(m_catename)).Read.Decimal(ss) = m_decimal

    ElseIf (FuseType = "SEN") Then
        SENFuse.Category(SENIndex(m_catename)).Read.Decimal(ss) = m_decimal
        
    ElseIf (FuseType = "MON") Then
        MONFuse.Category(MONIndex(m_catename)).Read.Decimal(ss) = m_decimal

    ElseIf (FuseType = "CMP") Then
        CMPFuse.Category(CMPIndex(m_catename)).Read.Decimal(ss) = m_decimal

    ElseIf (FuseType = "UDRE") Then
        UDRE_Fuse.Category(UDRE_Index(m_catename)).Read.Decimal(ss) = m_decimal

    ElseIf (FuseType = "UDRP") Then
        UDRP_Fuse.Category(UDRP_Index(m_catename)).Read.Decimal(ss) = m_decimal

    ElseIf (FuseType = "CMPE") Then
        CMPE_Fuse.Category(CMPE_Index(m_catename)).Read.Decimal(ss) = m_decimal

    ElseIf (FuseType = "CMPP") Then
        CMPP_Fuse.Category(CMPP_Index(m_catename)).Read.Decimal(ss) = m_decimal

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,MON,SEN,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If

    auto_eFuse_SetReadDecimal = m_decimal

    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + FormatNumeric("Fuse SetReadDecimal", -25)
        m_dlogstr = vbTab & "Site(" + CStr(ss) + ") " + FuseType + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_decimal, -10)
        TheExec.Datalog.WriteComment m_dlogstr
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150625 New Function
Public Function auto_eFuse_GetReadDecimal(ByVal FuseType As String, m_catename As String, Optional showPrint As Boolean = True) As Long
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_GetReadDecimal"

    Dim m_len As Long
    Dim m_decimal As Long
    Dim m_dlogstr As String
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber

    FuseType = UCase(Trim(FuseType))
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)

    If (FuseType = "ECID") Then
        m_decimal = ECIDFuse.Category(ECIDIndex(m_catename)).Read.Decimal(ss)

    ElseIf (FuseType = "CFG") Then
        m_decimal = CFGFuse.Category(CFGIndex(m_catename)).Read.Decimal(ss)

    ElseIf (FuseType = "UID") Then
        m_decimal = UIDFuse.Category(UIDIndex(m_catename)).Read.Decimal(ss)

    ElseIf (FuseType = "UDR") Then
        m_decimal = UDRFuse.Category(UDRIndex(m_catename)).Read.Decimal(ss)

    ElseIf (FuseType = "SEN") Then
        m_decimal = SENFuse.Category(SENIndex(m_catename)).Read.Decimal(ss)

    ElseIf (FuseType = "MON") Then
        m_decimal = MONFuse.Category(MONIndex(m_catename)).Read.Decimal(ss)

    ElseIf (FuseType = "CMP") Then
        m_decimal = CMPFuse.Category(CMPIndex(m_catename)).Read.Decimal(ss)

    ElseIf (FuseType = "UDRE") Then
        m_decimal = UDRE_Fuse.Category(UDRE_Index(m_catename)).Read.Decimal(ss)

    ElseIf (FuseType = "UDRP") Then
        m_decimal = UDRP_Fuse.Category(UDRP_Index(m_catename)).Read.Decimal(ss)

    ElseIf (FuseType = "CMPE") Then
        m_decimal = CMPE_Fuse.Category(CMPE_Index(m_catename)).Read.Decimal(ss)

    ElseIf (FuseType = "CMPP") Then
        m_decimal = CMPP_Fuse.Category(CMPP_Index(m_catename)).Read.Decimal(ss)

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,MON,SEN,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If
    
    auto_eFuse_GetReadDecimal = m_decimal
    
    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + FormatNumeric("Fuse GetReadDecimal", -25)
        m_dlogstr = vbTab & "Site(" + CStr(ss) + ") " + FuseType + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_decimal, -10)
        TheExec.Datalog.WriteComment m_dlogstr
    End If
       
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20161115 Add from Tcay and update
''''This function is used for Vddbin/IDS
''''20180522 update As Double to As Variant, m_decimal to Variant
Public Function auto_eFuse_GetReadValue(ByVal FuseType As String, m_catename As String, Optional showPrint As Boolean = True) As Variant
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_GetReadValue"

    Dim m_len As Long
    Dim m_decimal As Variant
    Dim m_dlogstr As String
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber

    FuseType = UCase(FuseType)
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)

    If (FuseType = "ECID") Then
        m_decimal = ECIDFuse.Category(ECIDIndex(m_catename)).Read.Value(ss)

    ElseIf (FuseType = "CFG") Then
        m_decimal = CFGFuse.Category(CFGIndex(m_catename)).Read.Value(ss)

    ElseIf ((FuseType = "AES") Or (FuseType = "UID")) Then
        m_decimal = UIDFuse.Category(UIDIndex(m_catename)).Read.Value(ss)

    ElseIf (FuseType = "UDR") Then
        m_decimal = UDRFuse.Category(UDRIndex(m_catename)).Read.Value(ss)

    ElseIf (FuseType = "SEN") Then
        m_decimal = SENFuse.Category(SENIndex(m_catename)).Read.Value(ss)

    ElseIf (FuseType = "MON") Then
        m_decimal = MONFuse.Category(MONIndex(m_catename)).Read.Value(ss)

    ElseIf (FuseType = "UDRE") Then
        m_decimal = UDRE_Fuse.Category(UDRE_Index(m_catename)).Read.Value(ss)

    ElseIf (FuseType = "UDRP") Then
        m_decimal = UDRP_Fuse.Category(UDRP_Index(m_catename)).Read.Value(ss)

    ElseIf (FuseType = "CMPE") Then
        m_decimal = CMPE_Fuse.Category(CMPE_Index(m_catename)).Read.Value(ss)

    ElseIf (FuseType = "CMPP") Then
        m_decimal = CMPP_Fuse.Category(CMPP_Index(m_catename)).Read.Value(ss)

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,MON,SEN,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If
    
    auto_eFuse_GetReadValue = m_decimal
    
    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + "Fuse GetReadValue "
        m_dlogstr = vbTab & "Site(" + CStr(ss) + ") " + FuseType + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_decimal, -10)
        TheExec.Datalog.WriteComment m_dlogstr
    End If
       
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150625 New Function
Public Function auto_eFuse_GetIDSResolution(ByVal FuseType As String, m_catename As String, Optional showPrint As Boolean = False) As Double
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_GetIDSResolution"

    Dim m_len As Long
    Dim m_resolution As Double
    Dim m_dlogstr As String

    FuseType = UCase(Trim(FuseType))
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)

    If (FuseType = "ECID") Then
        m_resolution = ECIDFuse.Category(ECIDIndex(m_catename)).Resoultion

    ElseIf (FuseType = "CFG") Then
        m_resolution = CFGFuse.Category(CFGIndex(m_catename)).Resoultion

    ElseIf (FuseType = "UID") Then
        m_resolution = UIDFuse.Category(UIDIndex(m_catename)).Resoultion

    ElseIf (FuseType = "UDR") Then
        m_resolution = UDRFuse.Category(UDRIndex(m_catename)).Resoultion

    ElseIf (FuseType = "SEN") Then
        m_resolution = SENFuse.Category(SENIndex(m_catename)).Resoultion

    ElseIf (FuseType = "MON") Then
        m_resolution = MONFuse.Category(MONIndex(m_catename)).Resoultion

    ElseIf (FuseType = "CMP") Then
        m_resolution = CMPFuse.Category(CMPIndex(m_catename)).Resoultion

    ElseIf (FuseType = "UDRE") Then
        m_resolution = UDRE_Fuse.Category(UDRE_Index(m_catename)).Resoultion

    ElseIf (FuseType = "UDRP") Then
        m_resolution = UDRP_Fuse.Category(UDRP_Index(m_catename)).Resoultion

    ElseIf (FuseType = "CMPE") Then
        m_resolution = CMPE_Fuse.Category(CMPE_Index(m_catename)).Resoultion

    ElseIf (FuseType = "CMPP") Then
        m_resolution = CMPP_Fuse.Category(CMPP_Index(m_catename)).Resoultion

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,MON,SEN,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If
    
    auto_eFuse_GetIDSResolution = m_resolution
    
    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + FormatNumeric("Fuse GetIDSResolution", -25)
        m_dlogstr = vbTab & FuseType + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_resolution, -10)
        TheExec.Datalog.WriteComment m_dlogstr
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20151007 New Add
''''20151222 Change the name from 'auto_get_IDS_Decimal' to 'auto_calc_IDS_Decimal'
Public Function auto_calc_IDS_Decimal(m_value As Variant, m_resolution As Double, Optional showPrint = False) As Long
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_calc_IDS_Decimal"
    
    Dim m_str As String
    Dim m_dbl As Double
    Dim m_dbl_round As Long
    Dim m_decimal As Long
    
    m_dbl = (m_value / m_resolution)
    m_dbl_round = Round(m_dbl)

    ''follow T-si rule
    If ((m_dbl_round - m_dbl) >= 0) Then ''''MUST have
        m_decimal = m_dbl_round
    Else
        m_decimal = m_dbl_round + 1
    End If
'    If ((m_dbl_round - m_dbl) > 0.000000000000001) Then ''''MUST have
'        m_decimal = m_dbl_round - 1
'    Else
'        m_decimal = m_dbl_round
'    End If

'    If ((m_dbl_round - m_dbl) > 0.000000000000001) Then ''''MUST have
'        m_decimal = m_dbl_round - 1
'    Else
'        m_decimal = m_dbl_round
'    End If
    auto_calc_IDS_Decimal = m_decimal
    
    If (showPrint) Then
        m_str = CStr(m_value) + "/" + CStr(m_resolution)
        Debug.Print funcName + "......" + m_str + " = m_Dbl = " + CStr(m_dbl)
        Debug.Print funcName + "......" + m_str + " = m_Dbl_round = " + CStr(m_dbl_round)
        Debug.Print funcName + "......" + m_str + " = " + CStr(m_decimal)
        Debug.Print "---------------------------------------------------------------------" + vbCrLf
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150625 New Function
Public Function auto_eFuse_IDS_SetWriteDecimal(ByVal FuseType As String, m_catename As String, m_value As Variant, _
                                               Optional showPrint As Boolean = True, _
                                               Optional chkStage As Boolean = True) As Long
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_IDS_SetWriteDecimal"
    
    Dim m_len As Long
    Dim i As Long
    Dim m_decimal As Long
    Dim m_bitwidth As Long
    Dim m_resolution As Double
    Dim m_dlogstr As String
    Dim m_stage As String
    Dim m_valStr As String
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber

    ''''20150707 update
    ''If (auto_eFuse_chkStage(fusetype, m_catename) = 0) Then Exit Function
    ''''20150707 update
    If (chkStage = True) Then
        If (auto_eFuse_chkStage(FuseType, m_catename) = 0) Then
            m_decimal = 0
        End If
    End If

    FuseType = UCase(Trim(FuseType))
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)
    m_resolution = auto_eFuse_GetIDSResolution(FuseType, m_catename)
    m_resolution = m_resolution * 0.001 ''''<NOTICE> update unit from mA to A
    
    ''m_decimal = Int(m_value / m_resolution) ''was
    
    ''''--------------------------------------------------------------
    ''''20151007 update method to match the test plan range
    ''''EX: 204.60mA/0.2mA ==> code=1023
    ''''    204.59mA/0.2mA ==> code=1022
    ''''    .......
    ''''    .......
    ''''      0.20mA/0.2mA ==> code=1
    ''''      0.19mA/0.2mA ==> code=0
    ''''      0.00mA/0.2mA ==> code=0
    ''''--------------------------------------------------------------
    ''''If print calculation result: use the below code
    ''''m_decimal = auto_calc_IDS_Decimal(m_value, m_resolution, True)
    ''''--------------------------------------------------------------
    m_decimal = auto_calc_IDS_Decimal(m_value, m_resolution)
    
    If (m_decimal < 0) Then m_decimal = 0 ''''prevent for the negative value
    
    If (FuseType = "ECID") Then
        i = ECIDIndex(m_catename)
        m_bitwidth = ECIDFuse.Category(i).BitWidth
        If (m_decimal >= (2 ^ m_bitwidth)) Then
            m_decimal = 0
            TheExec.Datalog.WriteComment "<WARING> Real Value is over bitwidth of BitDefTable"
        End If
        ''''If (m_decimal >= (2 ^ m_bitwidth)) Then m_decimal = (2 ^ m_bitwidth) - 1
        ECIDFuse.Category(i).Write.Decimal(ss) = m_decimal

    ElseIf (FuseType = "CFG") Then
        i = CFGIndex(m_catename)
        m_bitwidth = CFGFuse.Category(i).BitWidth
        If (m_decimal >= (2 ^ m_bitwidth)) Then
            m_decimal = 0
            TheExec.Datalog.WriteComment "<WARING> Real Value is over bitwidth of BitDefTable"
        End If
        ''''If (m_decimal >= (2 ^ m_bitwidth)) Then m_decimal = (2 ^ m_bitwidth) - 1
        CFGFuse.Category(i).Write.Decimal(ss) = m_decimal
    
    ElseIf (FuseType = "UID") Then
        i = UIDIndex(m_catename)
        m_bitwidth = UIDFuse.Category(i).BitWidth
        If (m_decimal >= (2 ^ m_bitwidth)) Then
            m_decimal = 0
            TheExec.Datalog.WriteComment "<WARING> Real Value is over bitwidth of BitDefTable"
        End If
        ''''If (m_decimal >= (2 ^ m_bitwidth)) Then m_decimal = (2 ^ m_bitwidth) - 1
        UIDFuse.Category(i).Write.Decimal(ss) = m_decimal
        
    ElseIf (FuseType = "UDR") Then
        i = UDRIndex(m_catename)
        m_bitwidth = UDRFuse.Category(i).BitWidth
        If (m_decimal >= (2 ^ m_bitwidth)) Then
            m_decimal = 0
            TheExec.Datalog.WriteComment "<WARING> Real Value is over bitwidth of BitDefTable"
        End If
        ''''If (m_decimal >= (2 ^ m_bitwidth)) Then m_decimal = (2 ^ m_bitwidth) - 1
        UDRFuse.Category(i).Write.Decimal(ss) = m_decimal
        
    ElseIf (FuseType = "SEN") Then
        i = SENIndex(m_catename)
        m_bitwidth = SENFuse.Category(i).BitWidth
        If (m_decimal >= (2 ^ m_bitwidth)) Then
            m_decimal = 0
            TheExec.Datalog.WriteComment "<WARING> Real Value is over bitwidth of BitDefTable"
        End If
        ''''If (m_decimal >= (2 ^ m_bitwidth)) Then m_decimal = (2 ^ m_bitwidth) - 1
        SENFuse.Category(i).Write.Decimal(ss) = m_decimal
        
    ElseIf (FuseType = "MON") Then
        i = MONIndex(m_catename)
        m_bitwidth = MONFuse.Category(i).BitWidth
        If (m_decimal >= (2 ^ m_bitwidth)) Then
            m_decimal = 0
            TheExec.Datalog.WriteComment "<WARING> Real Value is over bitwidth of BitDefTable"
        End If
        ''''If (m_decimal >= (2 ^ m_bitwidth)) Then m_decimal = (2 ^ m_bitwidth) - 1
        MONFuse.Category(i).Write.Decimal(ss) = m_decimal

    ElseIf (FuseType = "CMP") Then
        i = CMPIndex(m_catename)
        m_bitwidth = CMPFuse.Category(i).BitWidth
        If (m_decimal >= (2 ^ m_bitwidth)) Then
            m_decimal = 0
            TheExec.Datalog.WriteComment "<WARING> Real Value is over bitwidth of BitDefTable"
        End If
        ''''If (m_decimal >= (2 ^ m_bitwidth)) Then m_decimal = (2 ^ m_bitwidth) - 1
        CMPFuse.Category(i).Write.Decimal(ss) = m_decimal

    ElseIf (FuseType = "UDRE") Then
        i = UDRE_Index(m_catename)
        m_bitwidth = UDRE_Fuse.Category(i).BitWidth
        If (m_decimal >= (2 ^ m_bitwidth)) Then
            m_decimal = 0
            TheExec.Datalog.WriteComment "<WARING> Real Value is over bitwidth of BitDefTable"
        End If
        ''''If (m_decimal >= (2 ^ m_bitwidth)) Then m_decimal = (2 ^ m_bitwidth) - 1
        UDRE_Fuse.Category(i).Write.Decimal(ss) = m_decimal

    ElseIf (FuseType = "UDRP") Then
        i = UDRP_Index(m_catename)
        m_bitwidth = UDRP_Fuse.Category(i).BitWidth
        If (m_decimal >= (2 ^ m_bitwidth)) Then
            m_decimal = 0
            TheExec.Datalog.WriteComment "<WARING> Real Value is over bitwidth of BitDefTable"
        End If
        ''''If (m_decimal >= (2 ^ m_bitwidth)) Then m_decimal = (2 ^ m_bitwidth) - 1
        UDRP_Fuse.Category(i).Write.Decimal(ss) = m_decimal

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP)"
        GoTo errHandler
        ''''nothing
    End If
    
    auto_eFuse_IDS_SetWriteDecimal = m_decimal
    
    If (showPrint) Then
        m_valStr = Format(m_value * 1000, "0.000000")
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + FormatNumeric("Fuse IDS_SetWriteDecimal ", -25)
        m_dlogstr = vbTab & "Site(" + CStr(ss) + ") " + FuseType + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_decimal, -10) + _
                    " (" + FormatNumeric(m_valStr + " mA", 12) + _
                    " / " + Format(m_resolution * 1000, "0.000000") + "mA)"
        TheExec.Datalog.WriteComment m_dlogstr
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150625 New Function
''''20150702 New Function, update with the stage judgement
Public Function auto_eFuse_GetAllPatTestPass_Flag(ByVal FuseType As String, Optional showPrint As Boolean = False) As Boolean
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_GetAllPatTestPass_Flag"
    
    Dim i As Long
    Dim m_len As Long
    Dim m_flag As Boolean
    Dim m_stage As String
    Dim m_dlogstr As String
    Dim m_catename As String
    Dim ss As Variant
    Dim m_default_or_real As String
    ss = TheExec.sites.SiteNumber
    
    FuseType = UCase(Trim(FuseType))
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)
    m_flag = True ''''initial
    
    If (FuseType = "ECID") Then
        For i = 0 To UBound(ECIDFuse.Category)
            m_stage = LCase(ECIDFuse.Category(i).Stage)
            m_default_or_real = LCase(ECIDFuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = ECIDFuse.Category(i).PatTestPass_Flag(ss)
                m_catename = ECIDFuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i

    ElseIf (FuseType = "CFG") Then
        For i = 0 To UBound(CFGFuse.Category)
            m_stage = LCase(CFGFuse.Category(i).Stage)
            m_default_or_real = LCase(CFGFuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = CFGFuse.Category(i).PatTestPass_Flag(ss)
                m_catename = CFGFuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i

    ElseIf (FuseType = "UID") Then
        For i = 0 To UBound(UIDFuse.Category)
            m_stage = LCase(UIDFuse.Category(i).Stage)
            m_default_or_real = LCase(UIDFuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = UIDFuse.Category(i).PatTestPass_Flag(ss)
                m_catename = UIDFuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i
        
    ElseIf (FuseType = "UDR") Then
        For i = 0 To UBound(UDRFuse.Category)
            m_stage = LCase(UDRFuse.Category(i).Stage)
            m_default_or_real = LCase(UDRFuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = UDRFuse.Category(i).PatTestPass_Flag(ss)
                m_catename = UDRFuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i
        
    ElseIf (FuseType = "SEN") Then
        For i = 0 To UBound(SENFuse.Category)
            m_stage = LCase(SENFuse.Category(i).Stage)
            m_default_or_real = LCase(SENFuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = SENFuse.Category(i).PatTestPass_Flag(ss)
                m_catename = SENFuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i
        
    ElseIf (FuseType = "MON") Then
        For i = 0 To UBound(MONFuse.Category)
            m_stage = LCase(MONFuse.Category(i).Stage)
            m_default_or_real = LCase(MONFuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = MONFuse.Category(i).PatTestPass_Flag(ss)
                m_catename = MONFuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i

    ElseIf (FuseType = "CMP") Then
        For i = 0 To UBound(CMPFuse.Category)
            m_stage = LCase(CMPFuse.Category(i).Stage)
            m_default_or_real = LCase(CMPFuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = CMPFuse.Category(i).PatTestPass_Flag(ss)
                m_catename = CMPFuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i

    ElseIf (FuseType = "UDRE") Then
        For i = 0 To UBound(UDRE_Fuse.Category)
            m_stage = LCase(UDRE_Fuse.Category(i).Stage)
            m_default_or_real = LCase(UDRE_Fuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = UDRE_Fuse.Category(i).PatTestPass_Flag(ss)
                m_catename = UDRE_Fuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i

    ElseIf (FuseType = "UDRP") Then
        For i = 0 To UBound(UDRP_Fuse.Category)
            m_stage = LCase(UDRP_Fuse.Category(i).Stage)
            m_default_or_real = LCase(UDRP_Fuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = UDRP_Fuse.Category(i).PatTestPass_Flag(ss)
                m_catename = UDRP_Fuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i

    ElseIf (FuseType = "CMPE") Then
        For i = 0 To UBound(CMPE_Fuse.Category)
            m_stage = LCase(CMPE_Fuse.Category(i).Stage)
            m_default_or_real = LCase(CMPE_Fuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = CMPE_Fuse.Category(i).PatTestPass_Flag(ss)
                m_catename = CMPE_Fuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i

    ElseIf (FuseType = "CMPP") Then
        For i = 0 To UBound(CMPP_Fuse.Category)
            m_stage = LCase(CMPP_Fuse.Category(i).Stage)
            m_default_or_real = LCase(CMPP_Fuse.Category(i).Default_Real)
            If (gS_JobName = m_stage) And (m_default_or_real = "real") Then
            'If (gS_JobName = m_stage) Then
                m_flag = CMPP_Fuse.Category(i).PatTestPass_Flag(ss)
                m_catename = CMPP_Fuse.Category(i).Name
                If (m_flag = False) Then Exit For
            End If
        Next i

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If
    
    auto_eFuse_GetAllPatTestPass_Flag = m_flag
    
    ''''20161129 print out the failed category
    If (m_flag = False) Then
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment funcName + ":: " + FormatNumeric(FuseType, 4) + "Fuse, " + m_catename + " PatTest Fail."
    End If
    
    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = "JobName=" + UCase(gS_JobName) + " " + FuseType + FormatNumeric("Fuse GetAllPatTestPass_Flag", -28)
        m_dlogstr = vbTab & "Site(" + CStr(ss) + ") " + FuseType + " = " + FormatNumeric(m_flag, -5)
        TheExec.Datalog.WriteComment m_dlogstr
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150625 New Function
Public Function auto_eFuse_SetPatTestPass_Flag(ByVal FuseType As String, m_catename As String, m_flag As Boolean, Optional showPrint As Boolean = True) As Boolean
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_SetPatTestPass_Flag"

    Dim m_len As Long
    Dim m_dlogstr As String
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber
    
    FuseType = UCase(Trim(FuseType))
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)
    
    If (FuseType = "ECID") Then
        ECIDFuse.Category(ECIDIndex(m_catename)).PatTestPass_Flag(ss) = m_flag

    ElseIf (FuseType = "CFG") Then
        CFGFuse.Category(CFGIndex(m_catename)).PatTestPass_Flag(ss) = m_flag

    ElseIf (FuseType = "UID") Then
        UIDFuse.Category(UIDIndex(m_catename)).PatTestPass_Flag(ss) = m_flag

    ElseIf (FuseType = "UDR") Then
        UDRFuse.Category(UDRIndex(m_catename)).PatTestPass_Flag(ss) = m_flag

    ElseIf (FuseType = "SEN") Then
        SENFuse.Category(SENIndex(m_catename)).PatTestPass_Flag(ss) = m_flag
        
    ElseIf (FuseType = "MON") Then
        MONFuse.Category(MONIndex(m_catename)).PatTestPass_Flag(ss) = m_flag

    ElseIf (FuseType = "CMP") Then
        CMPFuse.Category(CMPIndex(m_catename)).PatTestPass_Flag(ss) = m_flag

    ElseIf (FuseType = "UDRE") Then
        UDRE_Fuse.Category(UDRE_Index(m_catename)).PatTestPass_Flag(ss) = m_flag

    ElseIf (FuseType = "UDRP") Then
        UDRP_Fuse.Category(UDRP_Index(m_catename)).PatTestPass_Flag(ss) = m_flag

    ElseIf (FuseType = "CMPE") Then
        CMPE_Fuse.Category(CMPE_Index(m_catename)).PatTestPass_Flag(ss) = m_flag

    ElseIf (FuseType = "CMPP") Then
        CMPP_Fuse.Category(CMPP_Index(m_catename)).PatTestPass_Flag(ss) = m_flag

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If
    
    auto_eFuse_SetPatTestPass_Flag = m_flag
    
    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = "Site(" + CStr(ss) + ") " + FuseType + FormatNumeric("Fuse SetPatTestPass_Flag", -25)
        m_dlogstr = vbTab & FormatNumeric(FuseType, Len(FuseType)) + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_flag, -5)
        TheExec.Datalog.WriteComment m_dlogstr
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
''''20150630 New Function
Public Function auto_eFuse_GetCatenameMaxLen(ByVal FuseType As String, Optional showPrint As Boolean = False) As Long
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_GetCatenameMaxLen"

    Dim m_len As Long
    Dim m_dlogstr As String

    FuseType = UCase(Trim(FuseType))

    If (FuseType = "ECID") Then
        m_len = gI_ECID_catename_maxLen

    ElseIf (FuseType = "CFG") Then
        m_len = gI_CFG_catename_maxLen

    ElseIf (FuseType = "UID") Then
        m_len = gI_UID_catename_maxLen

    ElseIf (FuseType = "UDR") Then
        m_len = gI_UDR_catename_maxLen

    ElseIf (FuseType = "SEN") Then
        m_len = gI_SEN_catename_maxLen
        
    ElseIf (FuseType = "MON") Then
        m_len = gI_MON_catename_maxLen

    ElseIf (FuseType = "CMP") Then
        m_len = gI_CMP_catename_maxLen

    ElseIf (FuseType = "UDRE") Then
        m_len = gI_UDRE_catename_maxLen

    ElseIf (FuseType = "UDRP") Then
        m_len = gI_UDRP_catename_maxLen

    ElseIf (FuseType = "CMPE") Then
        m_len = gI_CMPE_catename_maxLen

    ElseIf (FuseType = "CMPP") Then
        m_len = gI_CMPP_catename_maxLen

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If

    If (gL_eFuse_catename_maxLen = 0) Then
        auto_eFuse_GetCatenameMaxLen = m_len
    Else
        auto_eFuse_GetCatenameMaxLen = gL_eFuse_catename_maxLen  ''''20150702, to align the datalog once mixed fuse, ex, TMPS
    End If

    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + FormatNumeric("Fuse GetCatenameMaxLen", -25)
        m_dlogstr = vbTab & FuseType + "MaxLength = " + FormatNumeric(m_len, -5)
        TheExec.Datalog.WriteComment m_dlogstr
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150702 New, get Blank/FBC from startbit to endbit
Public Function auto_eFuse_GetBlankFBC_byBits(ByVal FuseType As String, SingleBitArray() As Long, m_startbit As Long, m_endbit As Long, ByRef blank As SiteBoolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_GetBlankFBC_byBits"
    
    Dim ss As Variant
    Dim k As Long, j As Long, i As Long
    Dim k1 As Long, k2 As Long
    Dim bcnt As Long
    Dim SingleSum As Long, DoubleSum As Long
    Dim TempDoubleBit As Long

    Dim BitsPerRow As Long
    Dim ReadCycles As Long
    Dim BitsPerCycle As Long
    Dim BitsPerBlock As Long
    Dim SingleDoubleFBC As Long
    
    SingleSum = 0: DoubleSum = 0: SingleDoubleFBC = 0
    ss = TheExec.sites.SiteNumber

    FuseType = UCase(Trim(FuseType))
    If (FuseType = "ECID") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = EcidBitsPerRow             ''''32   , 16  , 32
        ReadCycles = EcidReadCycle              ''''16   , 16  , 16
        BitsPerCycle = EcidReadBitWidth         ''''32   , 32  , 32
        BitsPerBlock = EcidBitPerBlockUsed      ''''256  , 256 , 512
        
    ElseIf (FuseType = "CFG") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = EConfigBitsPerRow          ''''32   , 16  , 32
        ReadCycles = EConfigReadCycle           ''''32   , 32  , 16
        BitsPerCycle = EConfigReadBitWidth      ''''32   , 32  , 32
        BitsPerBlock = EConfigBitPerBlockUsed   ''''512  , 512 , 512
    
    ElseIf (FuseType = "UID") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = UIDBitsPerRow              ''''32   , 16  , 32
        ReadCycles = UIDReadCycle               ''''64   , 64  , 32
        BitsPerCycle = UIDReadBitWidth          ''''32   , 32  , 32
        BitsPerBlock = UIDBitsPerBlockUsed      ''''1024 , 1024, 1024

    ElseIf (FuseType = "SEN") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = SENSORBitsPerRow           ''''32   , 16  , 32
        ReadCycles = SENSORReadCycle            ''''32   , 32  , 16
        BitsPerCycle = SENSORReadBitWidth       ''''32   , 32  , 32
        BitsPerBlock = SENSORBitPerBlockUsed    ''''512  , 512 , 512
        
    ElseIf (FuseType = "MON") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = MONITORBitsPerRow           ''''32   , 16  , 32
        ReadCycles = MONITORReadCycle            ''''32   , 32  , 16
        BitsPerCycle = MONITORReadBitWidth       ''''32   , 32  , 32
        BitsPerBlock = MONITORBitPerBlockUsed    ''''512  , 512 , 512

    ''''was FuseType = "UDR", 20171103 update for UDR,UDRE,UDRP
    ElseIf (FuseType Like "UDR*") Then
        ''''get the result in below
    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,SEN,MON,UDR,UDRE,UDRP)"
        GoTo errHandler
    End If

    ''''------------------------------------------------------------
    ''''20150720 update
    ''''UDR is the serial bits
    ''''was FuseType = "UDR", 20171103 update for UDR,UDRE,UDRP
    If (FuseType Like "UDR*") Then
        For k = m_startbit To m_endbit
            SingleSum = SingleSum + SingleBitArray(k)
            DoubleSum = DoubleSum + SingleBitArray(k)
            If (SingleBitArray(k) <> 0) Then
                blank(ss) = False
            End If
        Next k
        SingleDoubleFBC = DoubleSum - SingleSum

    ElseIf (gS_EFuse_Orientation = "UP2DOWN") Then

        For k = m_startbit To m_endbit
            SingleSum = SingleSum + (SingleBitArray(k) + SingleBitArray(BitsPerBlock + k))
            TempDoubleBit = SingleBitArray(k) Or SingleBitArray(BitsPerBlock + k)
            DoubleSum = DoubleSum + TempDoubleBit
            
            If (SingleBitArray(k) <> 0 Or SingleBitArray(BitsPerBlock + k) <> 0) Then
                blank(ss) = False
            End If
        Next k
        SingleDoubleFBC = DoubleSum * 2 - SingleSum

    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then

        bcnt = 0
        ''''total 32x16=512
        For k = 0 To ReadCycles - 1     '0 to 31
            For j = 0 To BitsPerRow - 1 '0 to 15
                ''''k1: Right block
                ''''k2:  Left block
                k1 = (k * BitsPerCycle) + j ''<Important> Must use XXXXReadBitWidth(=BitsPerCycle) here
                k2 = (k * BitsPerCycle) + BitsPerRow + j
                
                If (bcnt >= m_startbit And bcnt <= m_endbit) Then
                    SingleSum = SingleSum + (SingleBitArray(k1) + SingleBitArray(k2))
                    TempDoubleBit = SingleBitArray(k1) Or SingleBitArray(k2)
                    DoubleSum = DoubleSum + TempDoubleBit
                    If (SingleBitArray(k1) <> 0 Or SingleBitArray(k2) <> 0) Then
                        blank(ss) = False
                    End If
                Else
                    ''''over m_endbit bits, set k,j to up limit to escape for-loop
                    If (bcnt > m_endbit) Then
                        k = ReadCycles
                        j = BitsPerRow
                    End If
                End If
                bcnt = bcnt + 1
            Next j
        Next k
        SingleDoubleFBC = DoubleSum * 2 - SingleSum

    ElseIf (gS_EFuse_Orientation = "SingleUp") Then
    
        For k = m_startbit To m_endbit
            SingleSum = SingleSum + SingleBitArray(k)
            DoubleSum = DoubleSum + SingleBitArray(k)
            
            If (SingleBitArray(k) <> 0) Then
                blank(ss) = False
            End If
        Next k
        SingleDoubleFBC = DoubleSum - SingleSum
        
    End If
    ''''------------------------------------------------------------

    auto_eFuse_GetBlankFBC_byBits = SingleDoubleFBC

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150702 BlankChk/FBC by the programming Stage
''''20171103 update with UDR_E,UDR_P
Public Function auto_eFuse_BlankChk_FBC_byStage(ByVal FuseType As String, SingleBitArray() As Long, ByRef blank As SiteBoolean, Optional SingleDoubleFBC As Long)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_BlankChk_FBC_byStage"
    
    Dim ss As Variant
    Dim i As Long
    
    Dim m_stage As String
    Dim m_startbit As Long
    Dim m_endbit As Long
    Dim m_bitwidth As Long
    Dim blank2 As New SiteBoolean
    
    ss = TheExec.sites.SiteNumber
    blank(ss) = True
    SingleDoubleFBC = 0
    blank2(ss) = True

    FuseType = UCase(Trim(FuseType))
    If (FuseType = "ECID") Then
        For i = 0 To UBound(ECIDFuse.Category)
            m_stage = LCase(ECIDFuse.Category(i).Stage) ''''<Notice>
            If (gS_JobName = m_stage) Then
                m_startbit = ECIDFuse.Category(i).MSBbit
                m_endbit = ECIDFuse.Category(i).LSBbit
                SingleDoubleFBC = SingleDoubleFBC + auto_eFuse_GetBlankFBC_byBits(FuseType, SingleBitArray, m_startbit, m_endbit, blank)
            End If
        Next i

    ElseIf (FuseType = "CFG") Then
        For i = 0 To UBound(CFGFuse.Category)
            m_stage = LCase(CFGFuse.Category(i).Stage) ''''<Notice>
            If (gS_JobName = m_stage) Then ''''m_stage Like gS_JobName
                m_startbit = CFGFuse.Category(i).LSBbit
                m_endbit = CFGFuse.Category(i).MSBbit
                SingleDoubleFBC = SingleDoubleFBC + auto_eFuse_GetBlankFBC_byBits(FuseType, SingleBitArray, m_startbit, m_endbit, blank)
            End If
        Next i

    ElseIf (FuseType = "UID") Then
        For i = 0 To UBound(UIDFuse.Category)
            m_stage = LCase(UIDFuse.Category(i).Stage) ''''<Notice>
            If (gS_JobName = m_stage) Then
                m_startbit = UIDFuse.Category(i).LSBbit
                m_endbit = UIDFuse.Category(i).MSBbit
                SingleDoubleFBC = SingleDoubleFBC + auto_eFuse_GetBlankFBC_byBits(FuseType, SingleBitArray, m_startbit, m_endbit, blank)
            End If
        Next i

    ElseIf (FuseType = "SEN") Then
        For i = 0 To UBound(SENFuse.Category)
            m_stage = LCase(SENFuse.Category(i).Stage) ''''<Notice>
            If (gS_JobName = m_stage) Then
                m_startbit = SENFuse.Category(i).LSBbit
                m_endbit = SENFuse.Category(i).MSBbit
                SingleDoubleFBC = SingleDoubleFBC + auto_eFuse_GetBlankFBC_byBits(FuseType, SingleBitArray, m_startbit, m_endbit, blank)
            End If
        Next i

    ElseIf (FuseType = "MON") Then
        For i = 0 To UBound(MONFuse.Category)
            m_stage = LCase(MONFuse.Category(i).Stage) ''''<Notice>
            If (gS_JobName = m_stage) Then
                m_startbit = MONFuse.Category(i).LSBbit
                m_endbit = MONFuse.Category(i).MSBbit
                SingleDoubleFBC = SingleDoubleFBC + auto_eFuse_GetBlankFBC_byBits(FuseType, SingleBitArray, m_startbit, m_endbit, blank)
            End If
        Next i

    ElseIf (FuseType = "UDR") Then
        For i = 0 To UBound(UDRFuse.Category)
            m_stage = LCase(UDRFuse.Category(i).Stage) ''''<Notice>
            If (gS_JobName = m_stage) Then
                m_startbit = UDRFuse.Category(i).LSBbit
                m_endbit = UDRFuse.Category(i).MSBbit
                SingleDoubleFBC = SingleDoubleFBC + auto_eFuse_GetBlankFBC_byBits(FuseType, SingleBitArray, m_startbit, m_endbit, blank)
            End If
        Next i

    ElseIf (FuseType = "UDRE") Then
        For i = 0 To UBound(UDRE_Fuse.Category)
            m_stage = LCase(UDRE_Fuse.Category(i).Stage) ''''<Notice>
            If (gS_JobName = m_stage) Then
                m_startbit = UDRE_Fuse.Category(i).LSBbit
                m_endbit = UDRE_Fuse.Category(i).MSBbit
                SingleDoubleFBC = SingleDoubleFBC + auto_eFuse_GetBlankFBC_byBits(FuseType, SingleBitArray, m_startbit, m_endbit, blank)
            End If
        Next i

    ElseIf (FuseType = "UDRP") Then
        For i = 0 To UBound(UDRP_Fuse.Category)
            m_stage = LCase(UDRP_Fuse.Category(i).Stage) ''''<Notice>
            If (gS_JobName = m_stage) Then
                m_startbit = UDRP_Fuse.Category(i).LSBbit
                m_endbit = UDRP_Fuse.Category(i).MSBbit
                SingleDoubleFBC = SingleDoubleFBC + auto_eFuse_GetBlankFBC_byBits(FuseType, SingleBitArray, m_startbit, m_endbit, blank)
            End If
        Next i

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,SEN,MON,UDR,UDRE,UDRP)"
        GoTo errHandler
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150707 New Function
'''' auto_eFuse_chkStage = 1 ==> Job and Stage are Same.
'''' auto_eFuse_chkStage = 0 ==> Job and Stage are Different.
Public Function auto_eFuse_chkStage(ByVal FuseType As String, m_catename As String, Optional showPrint As Boolean = False) As Long
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_chkStage"

    Dim i As Long
    Dim m_stage As String
    Dim m_dlogstr As String
    Dim m_chkResult As Long
    
    FuseType = UCase(Trim(FuseType))

    If (FuseType = "ECID") Then
        i = ECIDIndex(m_catename)
        m_stage = LCase(ECIDFuse.Category(i).Stage)

    ElseIf (FuseType = "CFG") Then
        i = CFGIndex(m_catename)
        m_stage = LCase(CFGFuse.Category(i).Stage)

    ElseIf (FuseType = "UID") Then
        i = UIDIndex(m_catename)
        m_stage = LCase(UIDFuse.Category(i).Stage)

    ElseIf (FuseType = "UDR") Then
        i = UDRIndex(m_catename)
        m_stage = LCase(UDRFuse.Category(i).Stage)

    ElseIf (FuseType = "SEN") Then
        i = SENIndex(m_catename)
        m_stage = LCase(SENFuse.Category(i).Stage)
        
    ElseIf (FuseType = "MON") Then
        i = MONIndex(m_catename)
        m_stage = LCase(MONFuse.Category(i).Stage)

    ElseIf (FuseType = "CMP") Then
        i = CMPIndex(m_catename)
        m_stage = LCase(CMPFuse.Category(i).Stage)

    ElseIf (FuseType = "UDRE") Then
        i = UDRE_Index(m_catename)
        m_stage = LCase(UDRE_Fuse.Category(i).Stage)

    ElseIf (FuseType = "UDRP") Then
        i = UDRP_Index(m_catename)
        m_stage = LCase(UDRP_Fuse.Category(i).Stage)

    ElseIf (FuseType = "CMPE") Then
        i = CMPE_Index(m_catename)
        m_stage = LCase(CMPE_Fuse.Category(i).Stage)

    ElseIf (FuseType = "CMPP") Then
        i = CMPP_Index(m_catename)
        m_stage = LCase(CMPP_Fuse.Category(i).Stage)

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If

    If (m_stage = gS_JobName) Then
        m_chkResult = 1
    Else
        m_chkResult = 0
    End If
    auto_eFuse_chkStage = m_chkResult
    
    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + "Fuse chkStage (" + FormatNumeric(m_catename, gL_eFuse_catename_maxLen) + ") "
        m_dlogstr = vbTab & FuseType + " = " + FormatNumeric(m_chkResult, -5)
        TheExec.Datalog.WriteComment m_dlogstr
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20151229 New Function
'''' auto_eFuse_JobExistInStage = True ==> Job is one of the programming Stages.
'''' auto_eFuse_JobExistInStage = False==> Job is NOT one of the programming Stages.
Public Function auto_eFuse_JobExistInStage(ByVal FuseType As String, Optional showPrint As Boolean = False) As Boolean
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_JobExistInStage"

    Dim i As Long
    Dim m_stage As String
    Dim m_dlogstr As String
    Dim m_match_flag As Boolean

    FuseType = UCase(Trim(FuseType))
    m_match_flag = False

    If (FuseType = "ECID") Then
        For i = 0 To UBound(ECIDFuse.Category)
            m_stage = LCase(ECIDFuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
    ElseIf (FuseType = "CFG") Then
        For i = 0 To UBound(CFGFuse.Category)
            m_stage = LCase(CFGFuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
        ''''20170630 update, using A00 as Stage process
        If (gB_findCFGCondTable_flag And m_match_flag = False) Then
            For i = 0 To UBound(CFGTable.Category(0).condition)
                m_stage = LCase(CFGTable.Category(0).condition(i).Stage)
                If (gS_JobName = m_stage) Then
                    m_match_flag = True
                    Exit For
                End If
            Next i
        End If
    ElseIf (FuseType = "UID") Then
        For i = 0 To UBound(UIDFuse.Category)
            m_stage = LCase(UIDFuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
    ElseIf (FuseType = "UDR") Then
        For i = 0 To UBound(UDRFuse.Category)
            m_stage = LCase(UDRFuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
    ElseIf (FuseType = "SEN") Then
        For i = 0 To UBound(SENFuse.Category)
            m_stage = LCase(SENFuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
    ElseIf (FuseType = "MON") Then
        For i = 0 To UBound(MONFuse.Category)
            m_stage = LCase(MONFuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
    ElseIf (FuseType = "CMP") Then
        For i = 0 To UBound(CMPFuse.Category)
            m_stage = LCase(CMPFuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
    ElseIf (FuseType = "UDRE") Then
        For i = 0 To UBound(UDRE_Fuse.Category)
            m_stage = LCase(UDRE_Fuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
    ElseIf (FuseType = "UDRP") Then
        For i = 0 To UBound(UDRP_Fuse.Category)
            m_stage = LCase(UDRP_Fuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
    ElseIf (FuseType = "CMPE") Then
        For i = 0 To UBound(CMPE_Fuse.Category)
            m_stage = LCase(CMPE_Fuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
    ElseIf (FuseType = "CMPP") Then
        For i = 0 To UBound(CMPP_Fuse.Category)
            m_stage = LCase(CMPP_Fuse.Category(i).Stage)
            If (gS_JobName = m_stage) Then
                m_match_flag = True
                Exit For
            End If
        Next i
    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP,UDRPP)"
        GoTo errHandler
    End If

    auto_eFuse_JobExistInStage = m_match_flag

    If (showPrint And m_match_flag = False) Then ''''20160714 update, only show when un-match, i.e., m_match_flag=False
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + "Fuse " + FormatNumeric(UCase(gS_JobName), 6)
        If (auto_eFuse_JobExistInStage = True) Then m_dlogstr = vbTab & FuseType + " is Existed in the all Programming Stages."
        If (auto_eFuse_JobExistInStage = False) Then m_dlogstr = vbTab & FuseType + " is NOT existed in the all Programming Stages."
        TheExec.Datalog.WriteComment vbCrLf & m_dlogstr
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150807 New, getCompFBC (DoubleBitArray vs eFuse_Pgm_Bit) from startbit to endbit
Public Function auto_eFuse_GetCompFBC_byBits(ByVal FuseType As String, DoubleBitArray() As Long, eFuse_Pgm_Bit() As Long, m_startbit As Long, m_endbit As Long) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_GetCompFBC_byBits"
    
    ''Dim ss As Variant
    Dim k As Long, j As Long, i As Long
    Dim k1 As Long, k2 As Long
    Dim bcnt As Long

    Dim BitsPerRow As Long
    Dim RowPerBlock As Long
    Dim ReadBitWidth As Long
    Dim BitsPerBlockUsed As Long
    Dim CompFBC As Long
    
    CompFBC = 0
    ''ss = TheExec.Sites.SiteNumber

    FuseType = UCase(Trim(FuseType))
    If (FuseType = "ECID") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = EcidBitsPerRow             ''''32   , 16  , 32
        RowPerBlock = EcidRowPerBlock           ''''8    , 16  , 8
        ReadBitWidth = EcidReadBitWidth         ''''32   , 32  , 32

    ElseIf (FuseType = "CFG") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = EConfigBitsPerRow          ''''32   , 16  , 32
        RowPerBlock = EConfigRowPerBlock        ''''16   , 32  , 16
        ReadBitWidth = EConfigReadBitWidth      ''''32   , 32  , 32

    ElseIf (FuseType = "UID") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = UIDBitsPerRow              ''''32   , 16  , 32
        RowPerBlock = UIDRowPerBlock            ''''32   , 64  , 32
        ReadBitWidth = UIDReadBitWidth          ''''32   , 32  , 32

    ElseIf (FuseType = "SEN") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = SENSORBitsPerRow           ''''32   , 16  , 32
        RowPerBlock = SENSORRowPerBlock         ''''16   , 32  , 16
        ReadBitWidth = SENSORReadBitWidth       ''''32   , 32  , 32
        
    ElseIf (FuseType = "MON") Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = MONITORBitsPerRow           ''''32   , 16  , 32
        RowPerBlock = MONITORRowPerBlock         ''''16   , 32  , 16
        ReadBitWidth = MONITORReadBitWidth       ''''32   , 32  , 32

    ''''was FuseType = "UDR", 20171103 update for UDR,UDRE,UDRP
    ElseIf (FuseType Like "UDR*") Then
        ''''get the result in below
    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,MON,UDR)"
        GoTo errHandler
    End If

    ''''20151229 New
    BitsPerBlockUsed = RowPerBlock * BitsPerRow

    ''''------------------------------------------------------------
    ''''UDR is the serial bits
    ''''was FuseType = "UDR", 20171103 update for UDR,UDRE,UDRP
    If (FuseType Like "UDR*") Then
        For k = m_startbit To m_endbit
            If DoubleBitArray(k) <> eFuse_Pgm_Bit(k) Then
                CompFBC = CompFBC + 1
            End If
        Next k

    ElseIf (gS_EFuse_Orientation = "SingleUp") Then

        For k = m_startbit To m_endbit
            If DoubleBitArray(k) <> eFuse_Pgm_Bit(k) Then
                CompFBC = CompFBC + 1
            End If
        Next k

    ElseIf (gS_EFuse_Orientation = "UP2DOWN") Then
        
        For k = m_startbit To m_endbit
            ''''Up-Side
            If DoubleBitArray(k) <> eFuse_Pgm_Bit(k) Then
                CompFBC = CompFBC + 1
            End If
            ''''Down-Side
            If DoubleBitArray(k) <> eFuse_Pgm_Bit(k + BitsPerBlockUsed) Then
                CompFBC = CompFBC + 1
            End If
        Next k

    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then

        bcnt = 0
        For k = 0 To RowPerBlock - 1     ''0...31
            For j = 0 To BitsPerRow - 1  ''0...15, 16 bits per row
                If (bcnt >= m_startbit And bcnt <= m_endbit) Then
                    ''''Right-Side
                    If DoubleBitArray(k * BitsPerRow + j) <> eFuse_Pgm_Bit(k * ReadBitWidth + j) Then
                        CompFBC = CompFBC + 1
                    End If
                    ''''Left-Side
                    If DoubleBitArray(k * BitsPerRow + j) <> eFuse_Pgm_Bit(k * ReadBitWidth + BitsPerRow + j) Then
                        CompFBC = CompFBC + 1
                    End If
                Else
                    ''''over m_endbit bits, set k,j to up limit to escape for-loop
                    If (bcnt > m_endbit) Then
                        k = RowPerBlock
                        j = BitsPerRow
                    End If
                End If
                bcnt = bcnt + 1
            Next j
        Next k

    End If
    ''''------------------------------------------------------------

    auto_eFuse_GetCompFBC_byBits = CompFBC

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150807
''''By this way, it can be easy to compare DoubleBit and PgmBit in the specific categories
Public Function auto_eFuse_Compare_DoubleBit_PgmBit_byCategory(ByVal FuseType As String, catename_grp As String, DoubleBitArray() As Long, eFuse_Pgm_Bit() As Long, FailCnt As SiteLong)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Compare_DoubleBit_PgmBit_byCategory"
    
    Dim ss As Variant
    Dim i As Long, j As Long
    
    Dim m_stage As String
    Dim m_catename As String
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_bitwidth As Long
    Dim CompFBC As Long

    ss = TheExec.sites.SiteNumber
    CompFBC = 0
    FuseType = UCase(Trim(FuseType))

    ''''------------------------------------------------
    ''''Split catename_grp as String array
    ''''------------------------------------------------
    Dim m_catenameArr() As String
    Dim m_cateArr_elem As String
    Dim cateCNT As Long
    m_catenameArr = Split(Trim(catename_grp), ",")
    cateCNT = 0
    For j = 0 To UBound(m_catenameArr)
        m_cateArr_elem = Trim(m_catenameArr(j))
        If (m_cateArr_elem <> "") Then
            m_catenameArr(cateCNT) = m_cateArr_elem
            cateCNT = cateCNT + 1
        End If
    Next j
    If (cateCNT >= 1) Then
        ReDim Preserve m_catenameArr(cateCNT - 1)
    Else
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please assign Category Names (" + FuseType + ")."
        On Error GoTo errHandler
    End If
    ''''------------------------------------------------
   
    If (FuseType = "ECID") Then

        For j = 0 To UBound(m_catenameArr)
            m_cateArr_elem = UCase(m_catenameArr(j))
            For i = 0 To UBound(ECIDFuse.Category)
                m_catename = UCase(ECIDFuse.Category(i).Name)
                If (m_catename = m_cateArr_elem) Then
                    m_LSBbit = ECIDFuse.Category(i).LSBbit
                    m_MSBBit = ECIDFuse.Category(i).MSBbit
                    If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
                        ''''<Notice> In ECID fuse, m_MSBbit is the startbit and m_LSBbit is the endbit.
                        CompFBC = CompFBC + auto_eFuse_GetCompFBC_byBits(FuseType, DoubleBitArray, eFuse_Pgm_Bit, m_MSBBit, m_LSBbit)
                    Else
                        ''''20160118 New
                        ''''<Notice> In ECID fuse, m_LSBbit is the startbit and m_MSBbit is the endbit.
                        CompFBC = CompFBC + auto_eFuse_GetCompFBC_byBits(FuseType, DoubleBitArray, eFuse_Pgm_Bit, m_LSBbit, m_MSBBit)
                    End If
                End If
            Next i
        Next j

    ElseIf (FuseType = "CFG") Then

        For j = 0 To UBound(m_catenameArr)
            m_cateArr_elem = UCase(m_catenameArr(j))
            For i = 0 To UBound(CFGFuse.Category)
                m_catename = UCase(CFGFuse.Category(i).Name)
                If (m_catename = m_cateArr_elem) Then
                    m_LSBbit = CFGFuse.Category(i).LSBbit
                    m_MSBBit = CFGFuse.Category(i).MSBbit
                    CompFBC = CompFBC + auto_eFuse_GetCompFBC_byBits(FuseType, DoubleBitArray, eFuse_Pgm_Bit, m_LSBbit, m_MSBBit)
                End If
            Next i
        Next j

    ElseIf (FuseType = "UID") Then

        For j = 0 To UBound(m_catenameArr)
            m_cateArr_elem = UCase(m_catenameArr(j))
            For i = 0 To UBound(UIDFuse.Category)
                m_catename = UCase(UIDFuse.Category(i).Name)
                If (m_catename = m_cateArr_elem) Then
                    m_LSBbit = UIDFuse.Category(i).LSBbit
                    m_MSBBit = UIDFuse.Category(i).MSBbit
                    CompFBC = CompFBC + auto_eFuse_GetCompFBC_byBits(FuseType, DoubleBitArray, eFuse_Pgm_Bit, m_LSBbit, m_MSBBit)
                End If
            Next i
        Next j

    ElseIf (FuseType = "SEN") Then

        For j = 0 To UBound(m_catenameArr)
            m_cateArr_elem = UCase(m_catenameArr(j))
            For i = 0 To UBound(SENFuse.Category)
                m_catename = UCase(SENFuse.Category(i).Name)
                If (m_catename = m_cateArr_elem) Then
                    m_LSBbit = SENFuse.Category(i).LSBbit
                    m_MSBBit = SENFuse.Category(i).MSBbit
                    CompFBC = CompFBC + auto_eFuse_GetCompFBC_byBits(FuseType, DoubleBitArray, eFuse_Pgm_Bit, m_LSBbit, m_MSBBit)
                End If
            Next i
        Next j
        
    ElseIf (FuseType = "MON") Then

        For j = 0 To UBound(m_catenameArr)
            m_cateArr_elem = UCase(m_catenameArr(j))
            For i = 0 To UBound(MONFuse.Category)
                m_catename = UCase(MONFuse.Category(i).Name)
                If (m_catename = m_cateArr_elem) Then
                    m_LSBbit = MONFuse.Category(i).LSBbit
                    m_MSBBit = MONFuse.Category(i).MSBbit
                    CompFBC = CompFBC + auto_eFuse_GetCompFBC_byBits(FuseType, DoubleBitArray, eFuse_Pgm_Bit, m_LSBbit, m_MSBBit)
                End If
            Next i
        Next j

    ElseIf (FuseType = "UDR") Then

        For j = 0 To UBound(m_catenameArr)
            m_cateArr_elem = UCase(m_catenameArr(j))
            For i = 0 To UBound(UDRFuse.Category)
                m_catename = UCase(UDRFuse.Category(i).Name)
                If (m_catename = m_cateArr_elem) Then
                    m_LSBbit = UDRFuse.Category(i).LSBbit
                    m_MSBBit = UDRFuse.Category(i).MSBbit
                    CompFBC = CompFBC + auto_eFuse_GetCompFBC_byBits(FuseType, DoubleBitArray, eFuse_Pgm_Bit, m_LSBbit, m_MSBBit)
                End If
            Next i
        Next j

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,MON,UDR)"
        GoTo errHandler
    End If
    
    FailCnt(ss) = CompFBC
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150825 New Function to Replace auto_Dec_to_PgmArr_(CFG/UID/UDR/SEN)Result_byStage()
''''20160202 update code, add optional "chkStage"
''''20160608 update resval as Variant, was Long
Public Function auto_eFuse_Dec2PgmArr_Write_byStage(ByVal FuseType As String, p_stage As String, idx As Long, _
                                                    resval As Variant, LSBbit As Long, MSBbit As Long, _
                                                    ByRef PgmArr() As Long, Optional chkStage As Boolean = True) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Dec2PgmArr_Write_byStage"
    
    Dim i As Long
    Dim m_decimal As Variant
    Dim m_bitStrM As String
    Dim m_binarr() As Long
    Dim m_bitsum As Long
    Dim m_bitwidth As Long
    Dim ss As Variant
    Dim m_HexStr As String

    ss = TheExec.sites.SiteNumber

    ''''20160620 update
    If (auto_isHexString(CStr(resval))) Then
        If (auto_chkHexStr_isOver7FFFFFFF(CStr(resval)) = True) Then
            Call auto_eFuse_HexStr2PgmArr_Write_byStage(FuseType, p_stage, idx, resval, LSBbit, MSBbit, PgmArr(), chkStage)
            Exit Function ''''<MUST>
        Else
            resval = Replace(resval, "_", "") ''''20171103 update, for case like as 0x1234_ABCD_EFFE
            ''''20170911 update
            If (UCase(CStr(resval)) Like "0X*") Then
                resval = Replace(UCase(CStr(resval)), "0X", "", 1, 1)
            ElseIf (UCase(CStr(resval)) Like "X*") Then
                resval = Replace(UCase(CStr(resval)), "X", "", 1, 1)
            End If
            resval = CLng("&H" & CStr(resval)) ''''Here it's Hex2Dec
        End If
    ElseIf (auto_isBinaryString(CStr(resval))) Then ''''20171211 add
        Dim m_BinStr As String
        m_BinStr = Replace(UCase(CStr(resval)), "B", "", 1, 1) ''''remove the first "B" character
        m_BinStr = Replace(m_BinStr, "_", "")                  ''''for case like as b0001_0101
        
        ''''convert to Hex with the prefix "0x"
        m_HexStr = "0x" + auto_BinStr2HexStr(m_BinStr, 1)

        If (auto_chkHexStr_isOver7FFFFFFF(m_HexStr) = True) Then
            Call auto_eFuse_HexStr2PgmArr_Write_byStage(FuseType, p_stage, idx, m_HexStr, LSBbit, MSBbit, PgmArr(), chkStage)
            Exit Function ''''<MUST>
        Else
            ''''remove prefix 0x
            resval = Replace(UCase(m_HexStr), "0X", "", 1, 1)
            resval = CLng("&H" & resval) ''''Here it's Hex2Dec
        End If
    Else
        ''''20170911 update
        ''''not a hex String
        If (resval <= (CDbl(2 ^ 31) - 1)) Then ''''<= 0x7FFFFFFF
            ''''do Nothing
        Else
            ''''over Long range (> 0x7FFFFFFF)
            ''''Firstly, convert to Hex String with prefix '0x'
            m_HexStr = auto_Value2HexStr(resval)
            Call auto_eFuse_HexStr2PgmArr_Write_byStage(FuseType, p_stage, idx, m_HexStr, LSBbit, MSBbit, PgmArr(), chkStage)
            Exit Function ''''<MUST>
        End If
    End If
    
    ''''20160202 update
    If (chkStage = True) Then
        If (gS_JobName = LCase(p_stage)) Then
            m_decimal = resval
        Else
            ''''if Not this stage, force the PgmArr to zero value
            m_decimal = 0
        End If
    Else
        ''''in case, it will be in the simulation mode.
        m_decimal = resval

        ''''20160606 update for the simulation (Job < Programming Stage)
        If (TheExec.TesterMode = testModeOffline And checkJob_less_Stage_Sequence(p_stage) = True) Then
            m_decimal = 0
        End If
    End If

    ''''--------------------------------------------------------------------------------------------
    m_bitsum = 0
    m_bitwidth = Abs(MSBbit - LSBbit) + 1
    m_bitStrM = auto_Dec2Bin_EFuse(m_decimal, m_bitwidth, m_binarr)
    ''''20170630 update
    m_HexStr = "0x" + auto_BinStr2HexStr(m_bitStrM, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))

    ''''<Notice> m_binarr(0) is LSB
    If (LSBbit <= MSBbit) Then
        For i = 0 To UBound(m_binarr)
            PgmArr(LSBbit + i) = m_binarr(i)
            m_bitsum = m_bitsum + m_binarr(i)
        Next i
    Else
        ''''case:: LSBbit > MSBbit
        For i = 0 To UBound(m_binarr)
            PgmArr(LSBbit - i) = m_binarr(i)
            m_bitsum = m_bitsum + m_binarr(i)
        Next i
    End If

    auto_eFuse_Dec2PgmArr_Write_byStage = m_decimal
    
    ''''--------------------------------------------------------------------------------------------
    FuseType = UCase(Trim(FuseType))
    If (FuseType = "ECID") Then
        ''''Not Support,check it later
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (CFG,UID,UDR,SEN,MON)"
        GoTo errHandler

    ElseIf (FuseType = "CFG") Then
        CFGFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        CFGFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        CFGFuse.Category(idx).Write.Decimal(ss) = m_decimal
        CFGFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        CFGFuse.Category(idx).Write.Value(ss) = m_decimal
        CFGFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        CFGFuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "UID") Then
        UIDFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UIDFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UIDFuse.Category(idx).Write.Decimal(ss) = m_decimal
        UIDFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UIDFuse.Category(idx).Write.Value(ss) = m_decimal
        UIDFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UIDFuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "UDR") Then
        UDRFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UDRFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRFuse.Category(idx).Write.Decimal(ss) = m_decimal
        UDRFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UDRFuse.Category(idx).Write.Value(ss) = m_decimal
        UDRFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UDRFuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "SEN") Then
        SENFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        SENFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        SENFuse.Category(idx).Write.Decimal(ss) = m_decimal
        SENFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        SENFuse.Category(idx).Write.Value(ss) = m_decimal
        SENFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        SENFuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "MON") Then
        MONFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        MONFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        MONFuse.Category(idx).Write.Decimal(ss) = m_decimal
        MONFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        MONFuse.Category(idx).Write.Value(ss) = m_decimal
        MONFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        MONFuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "UDRE") Then
        UDRE_Fuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UDRE_Fuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRE_Fuse.Category(idx).Write.Decimal(ss) = m_decimal
        UDRE_Fuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UDRE_Fuse.Category(idx).Write.Value(ss) = m_decimal
        UDRE_Fuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UDRE_Fuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "UDRP") Then
        UDRP_Fuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UDRP_Fuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRP_Fuse.Category(idx).Write.Decimal(ss) = m_decimal
        UDRP_Fuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UDRP_Fuse.Category(idx).Write.Value(ss) = m_decimal
        UDRP_Fuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UDRP_Fuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (CFG,UID,UDR,SEN,MON,UDRE,UDRP)"
        GoTo errHandler
    End If
    ''''--------------------------------------------------------------------------------------------

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20150825 New Function to Replace auto_Dec_to_PgmArr_(CFG/UID/UDR/SEN/MON)Result_byCategory()
Public Function auto_eFuse_Dec2PgmArr_Write_byCategory(ByVal FuseType As String, cateflag As Boolean, idx As Long, resval As Variant, LSBbit As Long, MSBbit As Long, ByRef PgmArr() As Long) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Dec2PgmArr_Write_byCategory"
    
    Dim i As Long
    Dim m_decimal As Variant
    Dim m_bitStrM As String
    Dim m_binarr() As Long
    Dim m_bitsum As Long
    Dim m_bitwidth As Long
    Dim ss As Variant
    Dim m_HexStr As String

    ss = TheExec.sites.SiteNumber

    ''''20160620 update
    If (auto_isHexString(CStr(resval))) Then
        If (auto_chkHexStr_isOver7FFFFFFF(CStr(resval)) = True) Then
            Call auto_eFuse_HexStr2PgmArr_Write_byStage(FuseType, gS_JobName, idx, resval, LSBbit, MSBbit, PgmArr(), False)
            Exit Function ''''<MUST>
        Else
            resval = Replace(resval, "_", "") ''''20171103 update, for case like as 0x1234_ABCD_EFFE
            ''''20170911 update
            If (UCase(CStr(resval)) Like "0X*") Then
                resval = Replace(UCase(CStr(resval)), "0X", "", 1, 1)
            ElseIf (UCase(CStr(resval)) Like "X*") Then
                resval = Replace(UCase(CStr(resval)), "X", "", 1, 1)
            End If
            resval = CLng("&H" & CStr(resval)) ''''Here it's Hex2Dec
        End If
    Else
        ''''20170911 update
        ''''not a hex String
        If (resval <= (CDbl(2 ^ 31) - 1)) Then ''''<= 0x7FFFFFFF
            ''''do Nothing
        Else
            ''''over Long range (> 0x7FFFFFFF)
            ''''Firstly, convert to Hex String with prefix '0x'
            m_HexStr = auto_Value2HexStr(resval)
            Call auto_eFuse_HexStr2PgmArr_Write_byStage(FuseType, gS_JobName, idx, m_HexStr, LSBbit, MSBbit, PgmArr(), False)
            Exit Function ''''<MUST>
        End If
    End If

    If (cateflag) Then
        m_decimal = resval
    Else
        ''''if Not this Category, force the PgmArr to zero value
        m_decimal = 0
    End If
     
    ''''--------------------------------------------------------------------------------------------
    m_bitsum = 0
    m_bitwidth = Abs(MSBbit - LSBbit) + 1
    m_bitStrM = auto_Dec2Bin_EFuse(m_decimal, m_bitwidth, m_binarr)
    ''''20170630 update
    m_HexStr = "0x" + auto_BinStr2HexStr(m_bitStrM, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))

    ''''<Notice> m_binarr(0) is LSB
    If (LSBbit <= MSBbit) Then
        For i = 0 To UBound(m_binarr)
            PgmArr(LSBbit + i) = m_binarr(i)
            m_bitsum = m_bitsum + m_binarr(i)
        Next i
    Else
        ''''case:: LSBbit > MSBbit
        For i = 0 To UBound(m_binarr)
            PgmArr(LSBbit - i) = m_binarr(i)
            m_bitsum = m_bitsum + m_binarr(i)
        Next i
    End If

    auto_eFuse_Dec2PgmArr_Write_byCategory = m_decimal
    
    ''''--------------------------------------------------------------------------------------------
    FuseType = UCase(Trim(FuseType))
    If (FuseType = "ECID") Then
        ''''Not Support,check it later
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (CFG,UID,UDR,SEN)"
        GoTo errHandler

    ElseIf (FuseType = "CFG") Then
        CFGFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        CFGFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        CFGFuse.Category(idx).Write.Decimal(ss) = m_decimal
        CFGFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        CFGFuse.Category(idx).Write.Value(ss) = m_decimal
        CFGFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        CFGFuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "UID") Then
        UIDFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UIDFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UIDFuse.Category(idx).Write.Decimal(ss) = m_decimal
        UIDFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UIDFuse.Category(idx).Write.Value(ss) = m_decimal
        UIDFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UIDFuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "UDR") Then
        UDRFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UDRFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRFuse.Category(idx).Write.Decimal(ss) = m_decimal
        UDRFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UDRFuse.Category(idx).Write.Value(ss) = m_decimal
        UDRFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UDRFuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "SEN") Then
        SENFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        SENFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        SENFuse.Category(idx).Write.Decimal(ss) = m_decimal
        SENFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        SENFuse.Category(idx).Write.Value(ss) = m_decimal
        SENFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        SENFuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update
        
    ElseIf (FuseType = "MON") Then
        MONFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        MONFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        MONFuse.Category(idx).Write.Decimal(ss) = m_decimal
        MONFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        MONFuse.Category(idx).Write.Value(ss) = m_decimal
        MONFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        MONFuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "UDRE") Then
        UDRE_Fuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UDRE_Fuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRE_Fuse.Category(idx).Write.Decimal(ss) = m_decimal
        UDRE_Fuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UDRE_Fuse.Category(idx).Write.Value(ss) = m_decimal
        UDRE_Fuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UDRE_Fuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    ElseIf (FuseType = "UDRP") Then
        UDRP_Fuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UDRP_Fuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRP_Fuse.Category(idx).Write.Decimal(ss) = m_decimal
        UDRP_Fuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UDRP_Fuse.Category(idx).Write.Value(ss) = m_decimal
        UDRP_Fuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UDRP_Fuse.Category(idx).Write.HexStr(ss) = m_HexStr ''''20170630 update

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (CFG,UID,UDR,SEN,MON,UDRE,UDRP)"
        GoTo errHandler
    End If
    ''''--------------------------------------------------------------------------------------------

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20170810 update with Optional StartBit
Public Function auto_PrintAllPgmBits(PgmArray() As Long, ByVal TotalCycleNumber As Long, ByVal TotalBitNum As Long, ByVal BitNumPerRow As Long, Optional StartBit As Long = 0)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_PrintAllPgmBits"
    
    Dim Row_Str As String
    Dim i As Long, j As Long, k As Long
    Dim headerStr As String
    
    Dim ss As Variant
    
    ss = TheExec.sites.SiteNumber
    
    TheExec.Datalog.WriteComment ""
    
    Row_Str = ""
    
    headerStr = "====== Print Out EFuse Program Bits "
    headerStr = headerStr + "( " + TheExec.DataManager.instanceName + " )" + " (Site" + CStr(ss) + ")"
    headerStr = headerStr + "============"
    
    If (BitNumPerRow = 32 And (gS_EFuse_Orientation = "RIGHT2LEFT")) Then
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "           3322222222221111 1111110000000000"
        TheExec.Datalog.WriteComment "           1098765432109876 5432109876543210"
        TheExec.Datalog.WriteComment "           ---------------- ----------------"
    Else
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "           33222222222211111111110000000000"
        TheExec.Datalog.WriteComment "           10987654321098765432109876543210"
        TheExec.Datalog.WriteComment ""
    End If

    Dim m_flag As Boolean:: m_flag = True
    ''''k = 0
    k = StartBit
    If (BitNumPerRow = 32 And (gS_EFuse_Orientation = "RIGHT2LEFT")) Then
        For i = 0 To TotalCycleNumber - 1
            For j = 0 To BitNumPerRow - 1
                If (j = 16) Then
                    ''''ading one space " " between bit16 and bit15
                    Row_Str = CStr(PgmArray(k)) + " " + Row_Str
                Else
                    Row_Str = CStr(PgmArray(k)) + Row_Str
                End If
                k = k + 1
            Next j
        Next i
        ''''because adding one speace " " per row, so total bit number should add cycle numbers tp parse Row_Str
        BitNumPerRow = BitNumPerRow + 1
        TotalBitNum = TotalBitNum + TotalCycleNumber
    Else
'        Dim m_PerRowSize As Long:: m_PerRowSize = 32
'        For i = 0 To TotalCycleNumber / m_PerRowSize - 1
'            For j = 0 To m_PerRowSize - 1
'                Row_Str = CStr(PgmArray(k)) + Row_Str
'                k = k + 1
'            Next j
'        Next i
'        BitNumPerRow = m_PerRowSize
'        TotalCycleNumber = TotalCycleNumber / m_PerRowSize

        m_flag = False
        Dim m_PerRowSize As Long:: m_PerRowSize = 32
        TotalCycleNumber = TotalCycleNumber / m_PerRowSize
        If (TotalBitNum Mod m_PerRowSize <> 0) Then TotalCycleNumber = TotalCycleNumber + 1
        For i = 0 To TotalCycleNumber - 1
            For j = 0 To m_PerRowSize - 1
                'Row_Str = CStr(PgmArray(k)) + Row_Str
                If (k < TotalBitNum) Then
                    Row_Str = CStr(PgmArray(k)) + Row_Str
                Else
                    Row_Str = "0" + Row_Str
                End If
                k = k + 1
            Next j
        Next i
        BitNumPerRow = m_PerRowSize
    End If

    If (k = 0) Then
        For i = 1 To TotalCycleNumber
            TheExec.Datalog.WriteComment "Row = " & Format((i - 1), "000") & ": " & Mid(Row_Str, (TotalBitNum - (BitNumPerRow * i - 1)), BitNumPerRow)
        Next i
    Else
        Dim kk As Long
        kk = (StartBit / BitNumPerRow)
'        For i = 1 To TotalCycleNumber
'            TheExec.Datalog.WriteComment "Row = " & Format((kk + i - 1), "000") & ": " & Mid(Row_Str, (TotalBitNum - (BitNumPerRow * i - 1)), BitNumPerRow)
'        Next i
        If (m_flag) Then
            For i = 1 To TotalCycleNumber
                TheExec.Datalog.WriteComment "Row = " & Format((kk + i - 1), "000") & ": " & Mid(Row_Str, (TotalBitNum - (BitNumPerRow * i - 1)), BitNumPerRow)
                'TheExec.DataLog.WriteComment "Row = " & Format((kk + i - 1), "000") & ": " & Mid(Row_Str, (m_PerRowSize * TotalCycleNumber - (m_PerRowSize * i)) + 1, BitNumPerRow)
            Next i
        Else
            For i = 1 To TotalCycleNumber
               ' TheExec.DataLog.WriteComment "Row = " & Format((kk + i - 1), "000") & ": " & Mid(Row_Str, (TotalBitNum - (BitNumPerRow * i - 1)), BitNumPerRow)
                TheExec.Datalog.WriteComment "Row = " & Format((kk + i - 1), "000") & ": " & Mid(Row_Str, (m_PerRowSize * TotalCycleNumber - (m_PerRowSize * i)) + 1, BitNumPerRow)
            Next i
        End If
    End If

    TheExec.Datalog.WriteComment "====== End of printing out all Program Bit ============"
    TheExec.Datalog.WriteComment ""

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
      
End Function

Public Function auto_PrintAllBitbyDSSC(HramArray() As Long, _
                                       ByVal TotalCycleNumber As Long, _
                                       ByVal TotalBitNum As Long, _
                                       ByVal BitNumPerRow As Long, _
                                       Optional FuseTypeIsCFG As Boolean = False)
  
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_PrintAllBitbyDSSC"
    
    Dim Row_Str As String
    Dim L As Long, j As Long, k As Long, i As Long
    Dim max_expand As Long
    Dim K_divided As Long
    Dim Block As Long
    Dim headerStr As String
    Dim ss As Variant
    
    ss = TheExec.sites.SiteNumber
    
    If (gB_eFuse_newMethod And FuseTypeIsCFG) Then
        Dim cmpRes_Arr() As Long
        cmpRes_Arr = gDW_CFG_Read_cmpsgWavePerCyc(ss).Data
    End If
    
    TheExec.Datalog.WriteComment ""
    
    Row_Str = ""
    
    headerStr = "====== Efuse Data read from DSSC (Chip internal ORed data) "
    headerStr = headerStr + "( " + TheExec.DataManager.instanceName + " )" + " (Site" + CStr(ss) + ")"
    headerStr = headerStr + "============"

    If (BitNumPerRow = 8) Then
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "           00000000"
        TheExec.Datalog.WriteComment "           76543210"
        TheExec.Datalog.WriteComment ""
    ElseIf (BitNumPerRow = 16) Then
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "           1111110000000000"
        TheExec.Datalog.WriteComment "           5432109876543210"
        TheExec.Datalog.WriteComment ""
    ElseIf (BitNumPerRow = 32 And (gS_EFuse_Orientation = "RIGHT2LEFT")) Then
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "           3322222222221111 1111110000000000  C"
        TheExec.Datalog.WriteComment "           1098765432109876 5432109876543210  M"
        TheExec.Datalog.WriteComment "           ---------------- ----------------  P"
    
    ElseIf (BitNumPerRow = 32) Then
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "           33222222222211111111110000000000"
        TheExec.Datalog.WriteComment "           10987654321098765432109876543210"
        TheExec.Datalog.WriteComment ""
    Else
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "Wrong BitNumPerRow(" + CStr(BitNumPerRow) + "), only 8, 16 and 32 support"
        TheExec.Datalog.WriteComment "           33222222222211111111110000000000"
        TheExec.Datalog.WriteComment "           10987654321098765432109876543210"
        TheExec.Datalog.WriteComment ""
    End If

    k = 0
    If (BitNumPerRow = 32 And (gS_EFuse_Orientation = "RIGHT2LEFT")) Then
        For i = 0 To TotalCycleNumber - 1
            For j = 0 To BitNumPerRow - 1
                If (j = 16) Then
                    ''''ading one space " " between bit16 and bit15
                    Row_Str = CStr(HramArray(k)) + " " + Row_Str
                Else
                    Row_Str = CStr(HramArray(k)) + Row_Str
                End If
                k = k + 1
            Next j
        Next i
        ''''because adding one speace " " per row, so total bit number should add cycle numbers tp parse Row_Str
        BitNumPerRow = BitNumPerRow + 1
        TotalBitNum = TotalBitNum + TotalCycleNumber
    Else
        For i = 0 To TotalCycleNumber - 1
            For j = 0 To BitNumPerRow - 1
                Row_Str = CStr(HramArray(k)) + Row_Str
                k = k + 1
            Next j
        Next i
    End If

    If (gS_EFuse_Orientation = "RIGHT2LEFT" And gB_eFuse_newMethod = True) Then
        For i = 1 To TotalCycleNumber
            TheExec.Datalog.WriteComment "Row = " & Format((i - 1), "000") & ": " & Mid(Row_Str, (TotalBitNum - (BitNumPerRow * i - 1)), BitNumPerRow) & "  " & cmpRes_Arr(i - 1)
        Next i
    Else
        For i = 1 To TotalCycleNumber
            TheExec.Datalog.WriteComment "Row = " & Format((i - 1), "000") & ": " & Mid(Row_Str, (TotalBitNum - (BitNumPerRow * i - 1)), BitNumPerRow)
        Next i
    End If

    TheExec.Datalog.WriteComment "====== End of Efuse Data read from DSSC ============"
    TheExec.Datalog.WriteComment ""

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160108 New Function
'''' auto_eFuse_chkLoLimit
''''20171211 add, convert the Binary syntax to Hex string
Public Function auto_eFuse_chkLoLimit(ByVal FuseType As String, m_idx As Long, m_stage As String, ByRef m_lolmt As Variant, Optional showPrint As Boolean = False) As Variant
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_chkLoLimit"

    Dim i As Long
    Dim m_catename As String
    Dim m_dlogstr As String
    
    ''''20171211 add, convert the Binary syntax to Hex string
    If (auto_isBinaryString(CStr(m_lolmt))) Then ''''20171211 add
        Dim m_BinStr As String
        m_BinStr = Replace(UCase(CStr(m_lolmt)), "B", "", 1, 1) ''''remove the first "B" character
        m_BinStr = Replace(m_BinStr, "_", "")                   ''''for case like as b0001_0101
        
        ''''convert to Hex with the prefix "0x"
        m_lolmt = "0x" + auto_BinStr2HexStr(m_BinStr, 1)
    End If
    
    ''''20160217
    If (gB_eFuse_Disable_ChkLMT_Flag = True) Then
        Exit Function
    End If

    m_stage = LCase(m_stage)
    FuseType = UCase(Trim(FuseType))
    auto_eFuse_chkLoLimit = 0#  ''''Initialize and default
    
    ''''20160324 update
    If (checkJob_less_Stage_Sequence(m_stage) = True) Then
        m_lolmt = 0#
    End If

    ''''--------------------------------------------------------------
    ''''<NOTICE>
    ''''Production Flow Sequence Should be like as
    ''''   Case1:: "CP1->CP2-(CP3)->WLFT->FT1->FT2->FT3"
    ''''   Case2:: "CP1->CP2-(CP3)->FT1->FT2->FT3"
    ''''--------------------------------------------------------------
    
    ''''<Important>
    auto_eFuse_chkLoLimit = m_lolmt

    If (showPrint) Then
        If (FuseType = "ECID") Then
            m_catename = ECIDFuse.Category(m_idx).Name
        ElseIf (FuseType = "CFG") Then
            m_catename = CFGFuse.Category(m_idx).Name
        ElseIf (FuseType = "UID") Then
            m_catename = UIDFuse.Category(m_idx).Name
        ElseIf (FuseType = "UDR") Then
            m_catename = UDRFuse.Category(m_idx).Name
        ElseIf (FuseType = "SEN") Then
            m_catename = SENFuse.Category(m_idx).Name
        ElseIf (FuseType = "MON") Then
            m_catename = MONFuse.Category(m_idx).Name
        ElseIf (FuseType = "UDRE") Then
            m_catename = UDRE_Fuse.Category(m_idx).Name
        ElseIf (FuseType = "UDRP") Then
            m_catename = UDRP_Fuse.Category(m_idx).Name
        Else
            TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,UDRE,UDRP)"
            GoTo errHandler
        End If

        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + "Fuse chkLoLimit (" + FormatNumeric(m_catename, gL_eFuse_catename_maxLen) + ") "
        m_dlogstr = vbTab & FuseType + " = " + FormatNumeric(m_lolmt, -5)
        TheExec.Datalog.WriteComment m_dlogstr
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160108 New Function
'''' auto_eFuse_chkHiLimit
''''20171211 add, convert the Binary syntax to Hex string
Public Function auto_eFuse_chkHiLimit(ByVal FuseType As String, m_idx As Long, m_stage As String, ByRef m_hilmt As Variant, Optional showPrint As Boolean = False) As Variant
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_chkHiLimit"

    Dim i As Long
    Dim m_catename As String
    Dim m_dlogstr As String

    ''''20171211 add, convert the Binary syntax to Hex string
    If (auto_isBinaryString(CStr(m_hilmt))) Then ''''20171211 add
        Dim m_BinStr As String
        m_BinStr = Replace(UCase(CStr(m_hilmt)), "B", "", 1, 1) ''''remove the first "B" character
        m_BinStr = Replace(m_BinStr, "_", "")                   ''''for case like as b0001_0101
        
        ''''convert to Hex with the prefix "0x"
        m_hilmt = "0x" + auto_BinStr2HexStr(m_BinStr, 1)
    End If

    ''''20160217
    If (gB_eFuse_Disable_ChkLMT_Flag = True) Then
        Exit Function
    End If

    m_stage = LCase(m_stage)
    FuseType = UCase(Trim(FuseType))
    auto_eFuse_chkHiLimit = 0#  ''''Initialize and default
    
    ''''20160324 update
    If (checkJob_less_Stage_Sequence(m_stage) = True) Then
        m_hilmt = 0#
    End If

    ''''--------------------------------------------------------------
    ''''<NOTICE>
    ''''Production Flow Sequence Should be like as
    ''''   Case1:: "CP1->CP2-(CP3)->WLFT->FT1->FT2->FT3"
    ''''   Case2:: "CP1->CP2-(CP3)->FT1->FT2->FT3"
    ''''--------------------------------------------------------------
    
    ''''<Important>
    auto_eFuse_chkHiLimit = m_hilmt
    
    If (showPrint) Then
        If (FuseType = "ECID") Then
            m_catename = ECIDFuse.Category(m_idx).Name
        ElseIf (FuseType = "CFG") Then
            m_catename = CFGFuse.Category(m_idx).Name
        ElseIf (FuseType = "UID") Then
            m_catename = UIDFuse.Category(m_idx).Name
        ElseIf (FuseType = "UDR") Then
            m_catename = UDRFuse.Category(m_idx).Name
        ElseIf (FuseType = "SEN") Then
            m_catename = SENFuse.Category(m_idx).Name
        ElseIf (FuseType = "MON") Then
            m_catename = MONFuse.Category(m_idx).Name
        ElseIf (FuseType = "UDRE") Then
            m_catename = UDRE_Fuse.Category(m_idx).Name
        ElseIf (FuseType = "UDRP") Then
            m_catename = UDRP_Fuse.Category(m_idx).Name
        Else
            TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,UDRE,UDRP)"
            GoTo errHandler
        End If

        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + "Fuse chkHiLimit (" + FormatNumeric(m_catename, gL_eFuse_catename_maxLen) + ") "
        m_dlogstr = vbTab & FuseType + " = " + FormatNumeric(m_hilmt, -5)
        TheExec.Datalog.WriteComment m_dlogstr
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160323, used for the offline simulation
''''setblankbyStage = True, it means that all related categories are zero.
''''20161026 remove "Optional setblankbyStage As Boolean = False"
Public Function eFuseENGFakeValue_Sim()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "eFuseENGFakeValue_Sim"

    Dim ss As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim InstName As String
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_defreal As String
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_simValue As Variant
    Dim m_bitwidth As Long
    Dim m_stage As String
    Dim m_setWriteFlag As Boolean
    Dim m_levelIdx As Long
    Dim m_Pmode As Long ''''was m_vddbinEnum
    Dim m_resolution As Double
    Dim m_bitStrM As String
    Dim m_bitsum As Long
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_tmpStr As String
    Dim m_pwrpinNameBinCut As String
    Dim m_catenameIDS As String
    Dim m_resolutionIDS As Double
    Dim m_idsCurrent As Double
    Dim m_bincut_calcVoltage As Double
    Dim m_decimal As Long
    Dim m_BinCut_CPVmin As Double
    Dim m_BinCut_CPVmax As Double
    Dim m_BinCut_CPGB As Double
    Dim dummyArr() As Long
    ReDim dummyArr(EConfigTotalBitCount - 1)

    InstName = UCase(TheExec.DataManager.instanceName)
    TheExec.Datalog.WriteComment vbTab & "TestInstance:: " + TheExec.DataManager.instanceName
    TheExec.Datalog.WriteComment vbTab & ("******** Start of eFuseENGFakeValue_Sim **********")

    For Each ss In TheExec.sites

        If (InstName Like "*ECID*") Then
            For i = 0 To UBound(ECIDFuse.Category)
                m_stage = ECIDFuse.Category(i).Stage
                m_catename = ECIDFuse.Category(i).Name
                m_algorithm = LCase(ECIDFuse.Category(i).algorithm)
                m_defreal = LCase(ECIDFuse.Category(i).Default_Real)
                m_lolmt = ECIDFuse.Category(i).LoLMT
                m_hilmt = ECIDFuse.Category(i).HiLMT
                m_bitwidth = ECIDFuse.Category(i).BitWidth
                m_resolution = ECIDFuse.Category(i).Resoultion
                m_setWriteFlag = False ''''default
                m_simValue = 0
                If (m_algorithm <> "lotid" And m_algorithm <> "numeric" And m_algorithm <> "crc") Then ''''20161013 update
                    ''''20171225 update
                    m_lolmt = auto_HexStr2Value(auto_Value2HexStr(m_lolmt))
                    m_hilmt = auto_HexStr2Value(auto_Value2HexStr(m_hilmt))
                    If (m_algorithm = "ids") Then
                        ''m_simValue = 2 ^ Fix((0.3 + 0.6 * Rnd(1)) * m_bitwidth)
                        m_simValue = Fix((m_lolmt + (m_hilmt - m_lolmt) * (0.6 + 0.3 * Rnd(1))) / m_resolution) ''''20170811 update
                        m_setWriteFlag = True
                    ElseIf (m_defreal = "real") Then
                        m_simValue = auto_eFuse_GetWriteDecimal("ECID", m_catename, False) ''''20160623 update
                        If (m_simValue = 0) Then
                            If (UCase(m_catename) Like "*TMPS*TD*") Then ''''20160623 update
                                m_simValue = 302 + Fix(10 * Rnd(1))
                            Else
                                m_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.6 + 0.3 * Rnd(1)))
                                If (m_simValue > (2 ^ 30)) Then
                                    ''m_simValue = (2 ^ Fix((0.3 + 0.6 * Rnd(1)) * 30)) + ss
                                    'm_simValue = CDbl(2 ^ Fix((0.5 + 0.3 * Rnd(ss)) * m_bitwidth) + 2 ^ Fix((0.1 + 0.3 * Rnd(ss)) * m_bitwidth))
                                    m_simValue = Round(CDbl((m_lolmt + (m_hilmt - m_lolmt) * (0.6 + 0.3 * Rnd(ss)))))
                                End If
                            End If
                        End If
                        m_setWriteFlag = True
                    End If
                    If (m_setWriteFlag = True) Then
                        ''''201812XX mask
                        ''If (checkJob_less_Stage_Sequence(m_stage) = True) Then m_simValue = 0 ''''20160324 update
                        Call auto_eFuse_SetWriteDecimal("ECID", m_catename, m_simValue, True, False)
                    End If
                End If
            Next i
        End If

        If (InstName Like "*CFG*" Or InstName Like "*CONFIG*") Then
            For i = 0 To UBound(CFGFuse.Category)
                m_stage = CFGFuse.Category(i).Stage
                m_catename = CFGFuse.Category(i).Name
                m_algorithm = LCase(CFGFuse.Category(i).algorithm)
                m_defreal = LCase(CFGFuse.Category(i).Default_Real)
                m_lolmt = CFGFuse.Category(i).LoLMT
                m_hilmt = CFGFuse.Category(i).HiLMT
                m_bitwidth = CFGFuse.Category(i).BitWidth
                m_resolution = CFGFuse.Category(i).Resoultion
                m_setWriteFlag = False ''''default
                m_simValue = 0
'                If (m_algorithm = "vddbin") Then
'                    Debug.Print "Vddbin"
'                End If
                
                If (m_algorithm = "firstbits" Or m_algorithm = "cond") Then ''''20161107 update, 20170630 add "cond"
                    ''''doNothing

                ElseIf (m_algorithm <> "crc") Then ''''20161021 update for crc
                    If (m_defreal = "bincut") Then
                        m_simValue = auto_eFuse_GetWriteDecimal("CFG", m_catename, False) ''''20161121 update
                        If (m_simValue = 0) Then
                            m_Pmode = VddBinStr2Enum(m_catename)
                            m_pwrpinNameBinCut = UCase(AllBinCut(m_Pmode).powerPin)
                            m_levelIdx = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step
                            
                            m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(m_levelIdx) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(m_levelIdx)
                            m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
                            ''''was simulation
                            'm_lolmt = Fix((m_lolmt - gD_BaseVoltage) / m_resolution) + 1
                            'm_hilmt = Fix((m_hilmt - gD_BaseVoltage) / m_resolution) - 1
                            'm_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.3 + 0.6 * Rnd(1)))
    
                            ''''get the related IDS simulation value
                            For j = 0 To UBound(CFGFuse.Category)
                                m_algorithm = LCase(CFGFuse.Category(j).algorithm)
                                If (m_algorithm = "ids") Then
                                    m_catenameIDS = CFGFuse.Category(j).Name
                                    m_resolutionIDS = CFGFuse.Category(j).Resoultion
                                    If (m_catenameIDS Like "*" + m_pwrpinNameBinCut + "*") Then
                                        ''''here m_idscurrent is 'mA', m_resolution is "mA"
                                        m_decimal = auto_eFuse_GetWriteDecimal("CFG", m_catenameIDS, False)
                                        If (m_decimal = 0) Then
                                            m_decimal = auto_eFuse_GetReadDecimal("CFG", m_catenameIDS, False)
                                        End If
                                        m_idsCurrent = m_resolutionIDS * m_decimal
                                        Exit For
                                    End If
                                End If
                            Next j
                            
                            ''''get the related bicut simulation value
                            k = Floor(0.5 + m_levelIdx * Rnd(1))
                            m_BinCut_CPVmin = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(k)
                            m_BinCut_CPVmax = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(k)
                            m_BinCut_CPGB = BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(k)
                            m_lolmt = m_BinCut_CPVmin
                            m_hilmt = m_BinCut_CPVmax
                            
                            If (m_idsCurrent = 0#) Then m_idsCurrent = m_resolutionIDS * 3 ''''20170630 prevent Log(0)
    
                            m_bincut_calcVoltage = BinCut(m_Pmode, CurrentPassBinCutNum).c(k) - BinCut(m_Pmode, CurrentPassBinCutNum).M(k) * (Log(m_idsCurrent) / Log(10))
                            m_bincut_calcVoltage = m_resolution * Floor(m_bincut_calcVoltage / m_resolution)
                            If (m_bincut_calcVoltage < m_lolmt) Then
                                m_bincut_calcVoltage = m_lolmt + m_BinCut_CPGB
                            ElseIf (m_bincut_calcVoltage > m_hilmt) Then
                                m_bincut_calcVoltage = m_hilmt + m_BinCut_CPGB
                            Else
                                m_bincut_calcVoltage = m_bincut_calcVoltage + m_BinCut_CPGB
                            End If
                            m_simValue = CeilingValue((m_bincut_calcVoltage - gD_BaseVoltage) / m_resolution)
                            TheExec.Datalog.WriteComment Space(16) + m_catename + ", simValue=" & m_simValue & ", bincut_calcVoltage=" & m_bincut_calcVoltage
                        End If
                        m_setWriteFlag = True
                    ElseIf (m_defreal = "real") Then
                        ''''20171225 update
                        m_lolmt = auto_HexStr2Value(auto_Value2HexStr(m_lolmt))
                        m_hilmt = auto_HexStr2Value(auto_Value2HexStr(m_hilmt))
                        m_simValue = auto_eFuse_GetWriteDecimal("CFG", m_catename, False) ''''20160623 update
                        If (m_algorithm = "ids") Then
                            If (m_simValue = 0) Then
                                ''m_simValue = 2 ^ Fix((0.8 + 0.2 * Rnd(1)) * m_bitwidth) + 2 ^ Fix(0.9 * Rnd(1) * (m_bitwidth - 1))
                                m_simValue = Fix((m_lolmt + (m_hilmt - m_lolmt) * (0.6 + 0.3 * Rnd(1))) / m_resolution) ''''20170811 update
                            End If
                        ElseIf (m_simValue = 0) Then
                            'm_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(1)))
                            m_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.3 + 0.2 * Rnd(1)))
                            If (m_simValue > (2 ^ 30)) Then
                                ''m_simValue = 2 ^ Fix((0.6 + 0.3 * Rnd(1)) * 30) + ss
                                ''m_simValue = CDbl(2 ^ Fix((0.9 + 0.1 * Rnd(ss)) * m_bitwidth))
                                'm_simValue = Round(CDbl((m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(ss)))))
                                m_simValue = Round(CDbl((m_lolmt + (m_hilmt - m_lolmt) * (0.3 + 0.2 * Rnd(ss)))))
                            End If
                        End If
                        m_setWriteFlag = True
                    End If
                    If (m_setWriteFlag = True) Then
                        ''''201812XX mask
                        ''If (checkJob_less_Stage_Sequence(m_stage) = True) Then m_simValue = 0 ''''20160324 update
                        ''''20160623, it will cause zero Write buffer if HIP/TMPS already simulate setWriteDecimal()
                        ''If (setblankbyStage = True And LCase(m_stage) = gS_JobName) Then m_simValue = 0 ''''20160526 update
                        If (InstName Like "*WRITE*") Then ''''20160526 update
                            If (LCase(m_stage) = gS_JobName) Then
                                Call auto_eFuse_SetWriteDecimal("CFG", m_catename, m_simValue, True, False)
                            Else
                                Call auto_eFuse_SetWriteDecimal("CFG", m_catename, m_simValue, False, False)
                            End If
                        Else
                            ''''blank instance, only showPrint on these category (Job>m_stage)
                            If (checkJob_less_Stage_Sequence(m_stage) = False And gS_JobName <> LCase(m_stage)) Then ''''Job > m_stage
                                Call auto_eFuse_SetWriteDecimal("CFG", m_catename, m_simValue, True, False)
                            Else
                                ''''(Job<=m_stage) set showPrint=False, 20161107 update
                                Call auto_eFuse_SetWriteDecimal("CFG", m_catename, m_simValue, False, False)
                            End If
                        End If
                    End If
                End If
            Next i
        End If

        ''''20171103 add
        If (InstName Like "*UDRE*") Then
            For i = 0 To UBound(UDRE_Fuse.Category)
                m_stage = UDRE_Fuse.Category(i).Stage
                m_catename = UDRE_Fuse.Category(i).Name
                m_algorithm = LCase(UDRE_Fuse.Category(i).algorithm)
                m_defreal = LCase(UDRE_Fuse.Category(i).Default_Real)
                m_lolmt = UDRE_Fuse.Category(i).LoLMT
                m_hilmt = UDRE_Fuse.Category(i).HiLMT
                m_resolution = UDRE_Fuse.Category(i).Resoultion
                m_setWriteFlag = False ''''default
                m_simValue = 0

                If (True) Then
                    If (m_defreal = "bincut") Then
                        m_simValue = auto_eFuse_GetWriteDecimal("UDRE", m_catename, False) ''''20161121 update
                        If (m_simValue = 0) Then
                            m_Pmode = VddBinStr2Enum(m_catename)
                            m_pwrpinNameBinCut = UCase(AllBinCut(m_Pmode).powerPin)
                            m_levelIdx = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step
                            m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(m_levelIdx) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(m_levelIdx)
                            m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
    
                            ''''was simulation
                            ''m_lolmt = Fix((m_lolmt - gD_BaseVoltage) / m_resolution) + 1
                            ''m_hilmt = Fix((m_hilmt - gD_BaseVoltage) / m_resolution) - 1
                            ''m_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.1 + 0.2 * Rnd(1)))
    
                            ''''get the related IDS simulation value
                            For j = 0 To UBound(CFGFuse.Category)
                                m_algorithm = LCase(CFGFuse.Category(j).algorithm)
                                If (m_algorithm = "ids") Then
                                    m_catenameIDS = CFGFuse.Category(j).Name
                                    m_resolutionIDS = CFGFuse.Category(j).Resoultion
                                    If (m_catenameIDS Like "*" + m_pwrpinNameBinCut + "*") Then
                                        ''''here m_idscurrent is 'mA', m_resolution is "mA"
                                        m_decimal = auto_eFuse_GetWriteDecimal("CFG", m_catenameIDS, False)
                                        If (m_decimal = 0) Then
                                            m_decimal = auto_eFuse_GetReadDecimal("CFG", m_catenameIDS, False)
                                        End If
                                        m_idsCurrent = m_resolutionIDS * m_decimal
                                        Exit For
                                    End If
                                End If
                            Next j
    
                            ''''get the related bicut simulation value
                            k = Floor(0.5 + m_levelIdx * Rnd(1))
                            m_BinCut_CPVmin = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(k)
                            m_BinCut_CPVmax = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(k)
                            m_BinCut_CPGB = BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(k)
                            m_lolmt = m_BinCut_CPVmin
                            m_hilmt = m_BinCut_CPVmax
                            
                            If (m_idsCurrent = 0#) Then m_idsCurrent = m_resolutionIDS * 3 ''''20170630 prevent Log(0)
                            
                            m_bincut_calcVoltage = BinCut(m_Pmode, CurrentPassBinCutNum).c(k) - BinCut(m_Pmode, CurrentPassBinCutNum).M(k) * (Log(m_idsCurrent) / Log(10))
                            m_bincut_calcVoltage = m_resolution * Floor(m_bincut_calcVoltage / m_resolution)
                            If (m_bincut_calcVoltage < m_lolmt) Then
                                m_bincut_calcVoltage = m_lolmt + m_BinCut_CPGB
                            ElseIf (m_bincut_calcVoltage > m_hilmt) Then
                                m_bincut_calcVoltage = m_hilmt + m_BinCut_CPGB
                            Else
                                m_bincut_calcVoltage = m_bincut_calcVoltage + m_BinCut_CPGB
                            End If
                            m_simValue = CeilingValue((m_bincut_calcVoltage - gD_UDRE_BaseVoltage) / m_resolution)
                            If (m_simValue = -1#) Then
                                m_bincut_calcVoltage = m_bincut_calcVoltage + m_resolution * 10
                                m_simValue = CeilingValue((m_bincut_calcVoltage - gD_BaseVoltage) / m_resolution)
                            End If
                            TheExec.Datalog.WriteComment Space(16) + m_catename + ", simValue=" & m_simValue & ", bincut_calcVoltage=" & m_bincut_calcVoltage
                        End If
                        m_setWriteFlag = True
                    ElseIf (m_defreal = "real") Then
                        ''''20171225 update
                        m_lolmt = auto_HexStr2Value(auto_Value2HexStr(m_lolmt))
                        m_hilmt = auto_HexStr2Value(auto_Value2HexStr(m_hilmt))
                        m_simValue = auto_eFuse_GetWriteDecimal("UDRE", m_catename, False) ''''20160623 update
                        If (m_simValue = 0) Then
                            m_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(1)))
                            If (m_simValue > (2 ^ 30)) Then
                                ''m_simValue = 2 ^ Fix((0.3 + 0.6 * Rnd(1)) * 30) + ss
                                ''m_simValue = CDbl(2 ^ Fix((0.9 + 0.1 * Rnd(ss)) * m_bitwidth))
                                m_simValue = Round(CDbl((m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(ss)))))
                            End If
                        End If
                        m_setWriteFlag = True
                    End If
                End If
                If (m_setWriteFlag = True) Then
                    ''''201812XX mask
                    ''If (checkJob_less_Stage_Sequence(m_stage) = True) Then m_simValue = 0 ''''20160324 update
                    ''''20160623, it will cause zero Write buffer if HIP/TMPS already simulate setWriteDecimal()
                    ''If (setblankbyStage = True And LCase(m_stage) = gS_JobName) Then m_simValue = 0 ''''20160526 update
                    If (InstName Like "*UDRE*USI*") Then ''''20160526 update
                        If (LCase(m_stage) = gS_JobName) Then
                            Call auto_eFuse_SetWriteDecimal("UDRE", m_catename, m_simValue, True, False)
                        Else
                            Call auto_eFuse_SetWriteDecimal("UDRE", m_catename, m_simValue, False, False)
                        End If
                    Else
                        ''''blank(USO) instance, only showPrint on these category (Job>m_stage)
                        If (checkJob_less_Stage_Sequence(m_stage) = False And gS_JobName <> LCase(m_stage)) Then ''''Job > m_stage
                            Call auto_eFuse_SetWriteDecimal("UDRE", m_catename, m_simValue, True, False)
                        Else
                            ''''(Job<=m_stage) set showPrint=False, 20161107 update
                            Call auto_eFuse_SetWriteDecimal("UDRE", m_catename, m_simValue, False, False)
                        End If
                    End If
                End If
            Next i

        ''''20171103 add
        ElseIf (InstName Like "*UDRP*") Then
            For i = 0 To UBound(UDRP_Fuse.Category)
                m_stage = UDRP_Fuse.Category(i).Stage
                m_catename = UDRP_Fuse.Category(i).Name
                m_algorithm = LCase(UDRP_Fuse.Category(i).algorithm)
                m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
                m_lolmt = UDRP_Fuse.Category(i).LoLMT
                m_hilmt = UDRP_Fuse.Category(i).HiLMT
                m_resolution = UDRP_Fuse.Category(i).Resoultion
                m_setWriteFlag = False ''''default
                m_simValue = 0

                If (True) Then
                    If (m_defreal = "bincut") Then
                        m_simValue = auto_eFuse_GetWriteDecimal("UDRP", m_catename, False) ''''20161121 update
                        If (m_simValue = 0) Then
                            m_Pmode = VddBinStr2Enum(m_catename)
                            m_pwrpinNameBinCut = UCase(AllBinCut(m_Pmode).powerPin)
                            m_levelIdx = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step
                            m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(m_levelIdx) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(m_levelIdx)
                            m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
    
                            ''''was simulation
                            ''m_lolmt = Fix((m_lolmt - gD_BaseVoltage) / m_resolution) + 1
                            ''m_hilmt = Fix((m_hilmt - gD_BaseVoltage) / m_resolution) - 1
                            ''m_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.1 + 0.2 * Rnd(1)))
    
                            ''''get the related IDS simulation value
                            For j = 0 To UBound(CFGFuse.Category)
                                m_algorithm = LCase(CFGFuse.Category(j).algorithm)
                                If (m_algorithm = "ids") Then
                                    m_catenameIDS = CFGFuse.Category(j).Name
                                    m_resolutionIDS = CFGFuse.Category(j).Resoultion
                                    If (m_catenameIDS Like "*" + m_pwrpinNameBinCut + "*") Then
                                        ''''here m_idscurrent is 'mA', m_resolution is "mA"
                                        m_decimal = auto_eFuse_GetWriteDecimal("CFG", m_catenameIDS, False)
                                        If (m_decimal = 0) Then
                                            m_decimal = auto_eFuse_GetReadDecimal("CFG", m_catenameIDS, False)
                                        End If
                                        m_idsCurrent = m_resolutionIDS * m_decimal
                                        Exit For
                                    End If
                                End If
                            Next j
    
                            ''''get the related bicut simulation value
                            k = Floor(0.5 + m_levelIdx * Rnd(1))
                            m_BinCut_CPVmin = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(k)
                            m_BinCut_CPVmax = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(k)
                            m_BinCut_CPGB = BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(k)
                            m_lolmt = m_BinCut_CPVmin
                            m_hilmt = m_BinCut_CPVmax
                            
                            If (m_idsCurrent = 0#) Then m_idsCurrent = m_resolutionIDS * 3 ''''20170630 prevent Log(0)
                            
                            m_bincut_calcVoltage = BinCut(m_Pmode, CurrentPassBinCutNum).c(k) - BinCut(m_Pmode, CurrentPassBinCutNum).M(k) * (Log(m_idsCurrent) / Log(10))
                            m_bincut_calcVoltage = m_resolution * Floor(m_bincut_calcVoltage / m_resolution)
                            If (m_bincut_calcVoltage < m_lolmt) Then
                                m_bincut_calcVoltage = m_lolmt + m_BinCut_CPGB
                            ElseIf (m_bincut_calcVoltage > m_hilmt) Then
                                m_bincut_calcVoltage = m_hilmt + m_BinCut_CPGB
                            Else
                                m_bincut_calcVoltage = m_bincut_calcVoltage + m_BinCut_CPGB
                            End If
                            m_simValue = CeilingValue((m_bincut_calcVoltage - gD_UDRP_BaseVoltage) / m_resolution)
                            If (m_simValue = -1#) Then
                                m_bincut_calcVoltage = m_bincut_calcVoltage + m_resolution * 10
                                m_simValue = CeilingValue((m_bincut_calcVoltage - gD_BaseVoltage) / m_resolution)
                            End If
                            TheExec.Datalog.WriteComment Space(16) + m_catename + ", simValue=" & m_simValue & ", bincut_calcVoltage=" & m_bincut_calcVoltage
                        End If
                        m_setWriteFlag = True
                    ElseIf (m_defreal = "real") Then
                        ''''20171225 update
                        m_lolmt = auto_HexStr2Value(auto_Value2HexStr(m_lolmt))
                        m_hilmt = auto_HexStr2Value(auto_Value2HexStr(m_hilmt))
                        m_simValue = auto_eFuse_GetWriteDecimal("UDRP", m_catename, False) ''''20160623 update
                        If (m_simValue = 0) Then
                            m_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(1)))
                            If (m_simValue > (2 ^ 30)) Then
                                ''m_simValue = 2 ^ Fix((0.3 + 0.6 * Rnd(1)) * 30) + ss
                                ''m_simValue = CDbl(2 ^ Fix((0.9 + 0.1 * Rnd(ss)) * m_bitwidth))
                                m_simValue = Round(CDbl((m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(ss)))))
                            End If
                        End If
                        m_setWriteFlag = True
                    End If
                End If
                If (m_setWriteFlag = True) Then
                    ''''201812XX mask
                    ''If (checkJob_less_Stage_Sequence(m_stage) = True) Then m_simValue = 0 ''''20160324 update
                    ''''20160623, it will cause zero Write buffer if HIP/TMPS already simulate setWriteDecimal()
                    ''If (setblankbyStage = True And LCase(m_stage) = gS_JobName) Then m_simValue = 0 ''''20160526 update
                    If (InstName Like "*UDRP*USI*") Then ''''20160526 update
                        If (LCase(m_stage) = gS_JobName) Then
                            Call auto_eFuse_SetWriteDecimal("UDRP", m_catename, m_simValue, True, False)
                        Else
                            Call auto_eFuse_SetWriteDecimal("UDRP", m_catename, m_simValue, False, False)
                        End If
                    Else
                        ''''blank(USO) instance, only showPrint on these category (Job>m_stage)
                        If (checkJob_less_Stage_Sequence(m_stage) = False And gS_JobName <> LCase(m_stage)) Then ''''Job > m_stage
                            Call auto_eFuse_SetWriteDecimal("UDRP", m_catename, m_simValue, True, False)
                        Else
                            ''''(Job<=m_stage) set showPrint=False, 20161107 update
                            Call auto_eFuse_SetWriteDecimal("UDRP", m_catename, m_simValue, False, False)
                        End If
                    End If
                End If
            Next i
        ''''UDR case
        ElseIf (InstName Like "*UDR*") Then
            For i = 0 To UBound(UDRFuse.Category)
                m_stage = UDRFuse.Category(i).Stage
                m_catename = UDRFuse.Category(i).Name
                m_algorithm = LCase(UDRFuse.Category(i).algorithm)
                m_defreal = LCase(UDRFuse.Category(i).Default_Real)
                m_lolmt = UDRFuse.Category(i).LoLMT
                m_hilmt = UDRFuse.Category(i).HiLMT
                m_resolution = UDRFuse.Category(i).Resoultion
                m_setWriteFlag = False ''''default
                m_simValue = 0

                If (True) Then
                    If (m_defreal = "bincut") Then
                        m_simValue = auto_eFuse_GetWriteDecimal("UDR", m_catename, False) ''''20161121 update
                        If (m_simValue = 0) Then
                            m_Pmode = VddBinStr2Enum(m_catename)
                            m_pwrpinNameBinCut = UCase(AllBinCut(m_Pmode).powerPin)
                            m_levelIdx = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step
                            m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(m_levelIdx) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(m_levelIdx)
                            m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
    
                            ''''was simulation
                            ''m_lolmt = Fix((m_lolmt - gD_BaseVoltage) / m_resolution) + 1
                            ''m_hilmt = Fix((m_hilmt - gD_BaseVoltage) / m_resolution) - 1
                            ''m_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.1 + 0.2 * Rnd(1)))
    
                            ''''get the related IDS simulation value
                            For j = 0 To UBound(CFGFuse.Category)
                                m_algorithm = LCase(CFGFuse.Category(j).algorithm)
                                If (m_algorithm = "ids") Then
                                    m_catenameIDS = CFGFuse.Category(j).Name
                                    m_resolutionIDS = CFGFuse.Category(j).Resoultion
                                    If (m_catenameIDS Like "*" + m_pwrpinNameBinCut + "*") Then
                                        ''''here m_idscurrent is 'mA', m_resolution is "mA"
                                        m_decimal = auto_eFuse_GetWriteDecimal("CFG", m_catenameIDS, False)
                                        If (m_decimal = 0) Then
                                            m_decimal = auto_eFuse_GetReadDecimal("CFG", m_catenameIDS, False)
                                        End If
                                        m_idsCurrent = m_resolutionIDS * m_decimal
                                        Exit For
                                    End If
                                End If
                            Next j
    
                            ''''get the related bicut simulation value
                            k = Floor(0.5 + m_levelIdx * Rnd(1))
                            m_BinCut_CPVmin = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(k)
                            m_BinCut_CPVmax = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(k)
                            m_BinCut_CPGB = BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(k)
                            m_lolmt = m_BinCut_CPVmin
                            m_hilmt = m_BinCut_CPVmax
                            
                            If (m_idsCurrent = 0#) Then m_idsCurrent = m_resolutionIDS * 3 ''''20170630 prevent Log(0)
                            
                            m_bincut_calcVoltage = BinCut(m_Pmode, CurrentPassBinCutNum).c(k) - BinCut(m_Pmode, CurrentPassBinCutNum).M(k) * (Log(m_idsCurrent) / Log(10))
                            m_bincut_calcVoltage = m_resolution * Floor(m_bincut_calcVoltage / m_resolution)
                            If (m_bincut_calcVoltage < m_lolmt) Then
                                m_bincut_calcVoltage = m_lolmt + m_BinCut_CPGB
                            ElseIf (m_bincut_calcVoltage > m_hilmt) Then
                                m_bincut_calcVoltage = m_hilmt + m_BinCut_CPGB
                            Else
                                m_bincut_calcVoltage = m_bincut_calcVoltage + m_BinCut_CPGB
                            End If
                            m_simValue = CeilingValue((m_bincut_calcVoltage - gD_BaseVoltage) / m_resolution)
                            If (m_simValue = -1#) Then
                                m_bincut_calcVoltage = m_bincut_calcVoltage + m_resolution * 10
                                m_simValue = CeilingValue((m_bincut_calcVoltage - gD_BaseVoltage) / m_resolution)
                            End If
                            TheExec.Datalog.WriteComment Space(16) + m_catename + ", simValue=" & m_simValue & ", bincut_calcVoltage=" & m_bincut_calcVoltage
                        End If
                        m_setWriteFlag = True
                    ElseIf (m_defreal = "real") Then
                        ''''20171225 update
                        m_lolmt = auto_HexStr2Value(auto_Value2HexStr(m_lolmt))
                        m_hilmt = auto_HexStr2Value(auto_Value2HexStr(m_hilmt))
                        m_simValue = auto_eFuse_GetWriteDecimal("UDR", m_catename, False) ''''20160623 update
                        If (m_simValue = 0) Then
                            m_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(1)))
                            If (m_simValue > (2 ^ 30)) Then
                                ''m_simValue = 2 ^ Fix((0.3 + 0.6 * Rnd(1)) * 30) + ss
                                ''m_simValue = CDbl(2 ^ Fix((0.9 + 0.1 * Rnd(ss)) * m_bitwidth))
                                m_simValue = Round(CDbl((m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(ss)))))
                            End If
                        End If
                        m_setWriteFlag = True
                    End If
                End If
                If (m_setWriteFlag = True) Then
                    ''''201812XX mask
                    ''If (checkJob_less_Stage_Sequence(m_stage) = True) Then m_simValue = 0 ''''20160324 update
                    ''''20160623, it will cause zero Write buffer if HIP/TMPS already simulate setWriteDecimal()
                    ''If (setblankbyStage = True And LCase(m_stage) = gS_JobName) Then m_simValue = 0 ''''20160526 update
                    If (InstName Like "*USI*") Then ''''20160526 update
                        If (LCase(m_stage) = gS_JobName) Then
                            Call auto_eFuse_SetWriteDecimal("UDR", m_catename, m_simValue, True, False)
                        Else
                            Call auto_eFuse_SetWriteDecimal("UDR", m_catename, m_simValue, False, False)
                        End If
                    Else
                        ''''blank(USO) instance, only showPrint on these category (Job>m_stage)
                        If (checkJob_less_Stage_Sequence(m_stage) = False And gS_JobName <> LCase(m_stage)) Then ''''Job > m_stage
                            Call auto_eFuse_SetWriteDecimal("UDR", m_catename, m_simValue, True, False)
                        Else
                            ''''(Job<=m_stage) set showPrint=False, 20161107 update
                            Call auto_eFuse_SetWriteDecimal("UDR", m_catename, m_simValue, False, False)
                        End If
                    End If
                End If
            Next i
        End If
    
        If (InstName Like "*SENSOR*") Then
            For i = 0 To UBound(SENFuse.Category)
                m_stage = SENFuse.Category(i).Stage
                m_catename = SENFuse.Category(i).Name
                m_algorithm = LCase(SENFuse.Category(i).algorithm)
                m_defreal = LCase(SENFuse.Category(i).Default_Real)
                m_lolmt = SENFuse.Category(i).LoLMT
                m_hilmt = SENFuse.Category(i).HiLMT
                m_setWriteFlag = False ''''default
                m_simValue = 0
                If (m_algorithm <> "crc") Then
                    If (m_defreal = "real") Then
                        ''''20171225 update
                        m_lolmt = auto_HexStr2Value(auto_Value2HexStr(m_lolmt))
                        m_hilmt = auto_HexStr2Value(auto_Value2HexStr(m_hilmt))
                        m_simValue = auto_eFuse_GetWriteDecimal("SEN", m_catename, False) ''''20160623 update
                        If (m_simValue = 0) Then
                            m_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(1)))
                            If (m_simValue > (2 ^ 30)) Then
                                ''m_simValue = 2 ^ Fix((0.3 + 0.6 * Rnd(1)) * 30) + ss
                                ''m_simValue = CDbl(2 ^ Fix((0.9 + 0.1 * Rnd(ss)) * m_bitwidth))
                                m_simValue = Round(CDbl((m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(ss)))))
                            End If
                            ''''<NOTICE> 20160531 update, user maintain
                            If (UCase(m_catename) Like "*TRIMG*") Then
                                m_simValue = CFGFuse.Category(CFGIndex("TrimG_SOC_0")).Read.Decimal(ss)
                            ElseIf (UCase(m_catename) Like "*TRIMO*") Then
                                m_simValue = CFGFuse.Category(CFGIndex("TrimO_SOC_0")).Read.Decimal(ss)
                            End If
                        End If
                        m_setWriteFlag = True
                    End If
                    If (m_setWriteFlag = True) Then
                        ''''201812XX mask
                        ''If (checkJob_less_Stage_Sequence(m_stage) = True) Then m_simValue = 0 ''''20160324 update
                        ''''20160623, it will cause zero Write buffer if HIP/TMPS already simulate setWriteDecimal()
                        ''If (setblankbyStage = True And LCase(m_stage) = gS_JobName) Then m_simValue = 0 ''''20160526 update
                        If (InstName Like "*WRITE*") Then ''''20160526 update
                            If (LCase(m_stage) = gS_JobName) Then
                                Call auto_eFuse_SetWriteDecimal("SEN", m_catename, m_simValue, True, False)
                            Else
                                Call auto_eFuse_SetWriteDecimal("SEN", m_catename, m_simValue, False, False)
                            End If
                        Else
                            ''''blank instance, only showPrint on these category (Job>m_stage)
                            If (checkJob_less_Stage_Sequence(m_stage) = False And gS_JobName <> LCase(m_stage)) Then ''''Job > m_stage
                                Call auto_eFuse_SetWriteDecimal("SEN", m_catename, m_simValue, True, False)
                            Else
                                ''''(Job<=m_stage) set showPrint=False, 20161107 update
                                Call auto_eFuse_SetWriteDecimal("SEN", m_catename, m_simValue, False, False)
                            End If
                        End If
                    End If
                End If
            Next i
        End If

        If (InstName Like "*MON*" Or InstName Like "*MONITOR*") Then
            For i = 0 To UBound(MONFuse.Category)
                m_stage = MONFuse.Category(i).Stage
                m_catename = MONFuse.Category(i).Name
                m_algorithm = LCase(MONFuse.Category(i).algorithm)
                m_defreal = LCase(MONFuse.Category(i).Default_Real)
                m_lolmt = MONFuse.Category(i).LoLMT
                m_hilmt = MONFuse.Category(i).HiLMT
                m_setWriteFlag = False ''''default
                m_simValue = 0
                If (m_algorithm <> "crc") Then
                    If (m_defreal = "real") Then
                        ''''20171225 update
                        m_lolmt = auto_HexStr2Value(auto_Value2HexStr(m_lolmt))
                        m_hilmt = auto_HexStr2Value(auto_Value2HexStr(m_hilmt))
                        m_simValue = auto_eFuse_GetWriteDecimal("MON", m_catename, False) ''''20160623 update
                        If (m_simValue = 0) Then
                            m_simValue = Fix(m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(1)))
                            If (m_simValue > (2 ^ 30)) Then
                                ''m_simValue = 2 ^ Fix((0.3 + 0.6 * Rnd(1)) * 30) + ss
                                ''m_simValue = CDbl(2 ^ Fix((0.9 + 0.1 * Rnd(ss)) * m_bitwidth))
                                m_simValue = Round(CDbl((m_lolmt + (m_hilmt - m_lolmt) * (0.8 + 0.2 * Rnd(ss)))))
                            End If
                            ''''<NOTICE> 20160531 update, user maintain
                            If (UCase(m_catename) Like "*TRIMG*") Then
                                m_simValue = CFGFuse.Category(CFGIndex("TrimG_SOC_0")).Read.Decimal(ss)
                            ElseIf (UCase(m_catename) Like "*TRIMO*") Then
                                m_simValue = CFGFuse.Category(CFGIndex("TrimO_SOC_0")).Read.Decimal(ss)
                            End If
                        End If
                        m_setWriteFlag = True
                    End If
                    If (m_setWriteFlag = True) Then
                        ''''201812XX mask
                        ''If (checkJob_less_Stage_Sequence(m_stage) = True) Then m_simValue = 0 ''''20160324 update
                        ''''20160623, it will cause zero Write buffer if HIP/TMPS already simulate setWriteDecimal()
                        ''If (setblankbyStage = True And LCase(m_stage) = gS_JobName) Then m_simValue = 0 ''''20160526 update
                        If (InstName Like "*WRITE*") Then ''''20160526 update
                            If (LCase(m_stage) = gS_JobName) Then
                                Call auto_eFuse_SetWriteDecimal("MON", m_catename, m_simValue, True, False)
                            Else
                                Call auto_eFuse_SetWriteDecimal("MON", m_catename, m_simValue, False, False)
                            End If
                        Else
                            ''''blank instance, only showPrint on these category (Job>m_stage)
                            If (checkJob_less_Stage_Sequence(m_stage) = False And gS_JobName <> LCase(m_stage)) Then ''''Job > m_stage
                                Call auto_eFuse_SetWriteDecimal("MON", m_catename, m_simValue, True, False)
                            Else
                                ''''(Job<=m_stage) set showPrint=False, 20161107 update
                                Call auto_eFuse_SetWriteDecimal("MON", m_catename, m_simValue, False, False)
                            End If
                        End If
                    End If
                End If
            Next i
        End If

    Next ss

    TheExec.Datalog.WriteComment vbTab & ("******** End of eFuseENGFakeValue_Sim **********") & vbCrLf

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160324 New
''''20160831 update with FT1_25C and FT2_85C
Public Function checkJob_less_Stage_Sequence(stageName As String, Optional showPrint As Boolean = "False") As Boolean

On Error GoTo errHandler
    Dim funcName As String:: funcName = "checkJob_less_Stage_Sequence"

    Dim m_jobNum As Long
    Dim m_stageNum As Long

    ''''----------------------------------------------------------
    ''''standard production flow is
    '''' CP1-->CP2-->CP3-->WLFT-->FT1(FT1_25C)-->FT2(FT2_85C)-->FT3 (-->FT4-->FT5)
    ''''----------------------------------------------------------

    Select Case UCase(gS_JobName)
        Case "CP1_EARLY" ''''20171016 update
            m_jobNum = 0
        Case "CP1"
            m_jobNum = 1
        Case "CP2"
            m_jobNum = 2
        Case "CP3"
            m_jobNum = 3
        Case "WLFT", "WLFT1"
            m_jobNum = 10
        Case "FT1"
            m_jobNum = 11
        Case "FT1_25C"
            m_jobNum = 12
        Case "FT2"
            m_jobNum = 13
        Case "FT2_85C"
            m_jobNum = 14
        Case "FT3"
            m_jobNum = 15
        Case "FT4"
            m_jobNum = 16
        Case "FT5"
            m_jobNum = 17
        Case Else
            m_jobNum = 99
    End Select
   
    Select Case UCase(stageName)
        Case "CP1_EARLY" ''''20171016 update
            m_stageNum = 0
        Case "CP1"
            m_stageNum = 1
        Case "CP2"
            m_stageNum = 2
        Case "CP3"
            m_stageNum = 3
        Case "WLFT", "WLFT1"
            m_stageNum = 10
        Case "FT1"
            m_stageNum = 11
        Case "FT1_25C"
            m_stageNum = 12
        Case "FT2"
            m_stageNum = 13
        Case "FT2_85C"
            m_stageNum = 14
        Case "FT3"
            m_stageNum = 15
        Case "FT4"
            m_stageNum = 16
        Case "FT5"
            m_stageNum = 17
        Case Else
            m_stageNum = 99
    End Select

    If (m_stageNum > m_jobNum) Then
        ''''means that (Job < Stage) setWrite '0' for simulation
        ''''means that setLimit = 0
        checkJob_less_Stage_Sequence = True
    Else
        ''''means that (Job >= Stage) need to setWrite simulate
        ''''means that setLimit as sheet
        checkJob_less_Stage_Sequence = False
        showPrint = False
    End If

    If (showPrint = True) Then
        Dim m_tmpStr As String
        m_tmpStr = funcName + "=" + CStr(checkJob_less_Stage_Sequence)
        m_tmpStr = m_tmpStr + " :: Job = " + UCase(gS_JobName) + "(m_jobNum=" + CStr(m_jobNum) + ")"
        m_tmpStr = m_tmpStr + ", Pgm_Stage = " + UCase(stageName) + "(m_stageNum=" + CStr(m_stageNum) + ")"
        TheExec.Datalog.WriteComment m_tmpStr
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function



''''20170811 update
Public Function auto_isHexString(ByVal InputStr As String) As Boolean

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_isHexString"
    
    Dim i As Long, j As Long
    Dim m_len As Long
    Dim m_char As String
    Dim HexChar() As Variant
    Dim m_match_flag As Boolean

    InputStr = UCase(InputStr)
    
    HexChar = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F")

    m_match_flag = False ''''<MUST> initialize
    
    InputStr = Replace(InputStr, "_", "") ''''20171103 update, for case like as 0x1234_ABCD_EFFE

    ''''20170811 update
    If InputStr Like UCase("0X*") Then ''''case "0xABCD"
        InputStr = Replace(InputStr, "0X", "", 1, 1)
        m_match_flag = True
    ElseIf InputStr Like UCase("X*") Then ''''case "xABCD"
        InputStr = Replace(InputStr, "X", "", 1, 1)
        m_match_flag = True
    Else
        m_match_flag = False
        ''TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: It's NOT the HEX String format, should have prefix '0x' as Hex-String."
    End If

    ''''do the advanced analysis
    If (m_match_flag = True) Then
        m_len = Len(InputStr)
        For i = 1 To m_len
            m_match_flag = False ''''<MUST> initialize per character, 20160616 update
            m_char = Mid(InputStr, i, 1)
            For j = 0 To UBound(HexChar)
                If (m_char = CStr(HexChar(j))) Then
                    m_match_flag = True
                    Exit For
                End If
            Next j
            If (m_match_flag = False) Then Exit For ''''<NOTICE>
        Next i
    End If

    auto_isHexString = m_match_flag

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20160414 New.
''''20160707 update for the case which Bitwidth is NOT multiple of 4bits(Hex)
''''20170811 update
Public Function auto_HexStr2BinStr_EFUSE(ByVal InputStr As String, BitWidth As Long, ByRef binarr() As Long) As String
                                                                                                                         
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_HexStr2BinStr_EFUSE"

    Dim i As Long, j As Long
    Dim PerChar As String
    Dim PerLetter As String
    Dim BinStr As String   ''''here its [MSB...LSB]
    Dim BinStr_L As String ''''here its [LSB...MSB]
    Dim DecodeBin As String
    Dim MyArray() As Variant
    Dim myArrayBin() As Variant
    Dim m_inputStr As String
    Dim m_len As Long
    Dim m_dummy As Long
    
    ReDim binarr(BitWidth - 1)

    If (auto_isHexString(InputStr) = False) Then
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: It's NOT HexString with the prefix '0x', please check it out."
        Exit Function
    Else
        ''''20170811 update
        InputStr = UCase(InputStr)
        InputStr = Replace(InputStr, "_", "") ''''20171103 update, for case like as 0x1234_ABCD_EFFE
        If (InputStr Like "0X*") Then
            m_inputStr = Replace(UCase(InputStr), "0X", "", 1, 1)
        ElseIf (InputStr Like "X*") Then
            m_inputStr = Replace(UCase(InputStr), "X", "", 1, 1)
        End If
    End If

    MyArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F")
    myArrayBin = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", "1000", "1001", _
                       "1010", "1011", "1100", "1101", "1110", "1111")
    BinStr = ""
    For i = 1 To Len(m_inputStr)
        PerChar = Mid(m_inputStr, i, 1)
        'One-to-One mapping, myarray() mappping to myarraybin()
        For j = 0 To UBound(MyArray)
            If (PerChar = MyArray(j)) Then
               DecodeBin = myArrayBin(j)
               Exit For
            End If
        Next j
        BinStr = BinStr + DecodeBin
    Next i

    BinStr_L = StrReverse(BinStr) ''''is [LSB......MSB]
    m_len = Len(BinStr_L)

    ''''<Important> 20160707 update, 20180723 update with the case (BitWidth = 0)
    If (BitWidth = m_len Or BitWidth = 0) Then
        ''''do Nothing
    ElseIf (BitWidth > m_len) Then
        ''''compensate "0" in the MSB bits
        For i = 1 To (BitWidth - m_len) ''''20160930 update i=1
            BinStr_L = BinStr_L + "0"
        Next i
    ElseIf (BitWidth < m_len) Then
        ''''get rid of "0" in the MSB bits
        ''''Ex: BitWidth=10bits, BinStr_L="110100000000"(12bits) => update BinStr_L="1101000000"(10bits)
        m_dummy = CLng(Mid(BinStr_L, BitWidth + 1, m_len - BitWidth)) ''''<MUST> put before the next statement, otherwise BinStr_L was update.
        BinStr_L = Mid(BinStr_L, 1, BitWidth)
        If (m_dummy > 0) Then
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: bitwidth(" + CStr(BitWidth) + ") < (" + CStr(Len(BinStr_L)) + ") " + BinStr_L + " (m_dummy=" & m_dummy & ")"
            GoTo errHandler
        End If
    End If
    BinStr = StrReverse(BinStr_L) ''''20160930 update
    For i = 1 To Len(BinStr_L)
        binarr(i - 1) = CLng(Mid(BinStr_L, i, 1))
    Next i

    auto_HexStr2BinStr_EFUSE = BinStr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160414 New Function to Use HexString ('0x')as input for the case which bits over 32bits
Public Function auto_eFuse_HexStr2PgmArr_Write_byStage(ByVal FuseType As String, p_stage As String, idx As Long, _
                                                       resval As Variant, LSBbit As Long, MSBbit As Long, _
                                                       ByRef PgmArr() As Long, Optional chkStage As Boolean = True) As Variant

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_HexStr2PgmArr_Write_byStage"

    Dim i As Long
    Dim m_HexStr As String
    Dim m_bitStrM As String
    Dim m_binarr() As Long
    Dim m_bitsum As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant
    Dim ss As Variant

    ss = TheExec.sites.SiteNumber

    ''''20160202 update
    If (chkStage = True) Then
        If (gS_JobName = LCase(p_stage)) Then
            m_HexStr = CStr(resval)
        Else
            ''''if Not this stage, force the PgmArr to zero value
            m_HexStr = "0x0"
        End If
    Else
        ''''in case, it will be in the simulation mode.
        m_HexStr = CStr(resval)
    End If
    m_HexStr = Replace(m_HexStr, "_", "") ''''20171103 update, for case like as 0x1234_ABCD_EFFE
     
    ''''--------------------------------------------------------------------------------------------
    m_bitsum = 0
    m_bitwidth = Abs(MSBbit - LSBbit) + 1
    m_bitStrM = auto_HexStr2BinStr_EFUSE(m_HexStr, m_bitwidth, m_binarr)
    m_decimal = auto_HexStr2Value(m_HexStr) ''''20170911 update
    
    ''''<Notice> m_binarr(0) is LSB
    If (LSBbit <= MSBbit) Then
        For i = 0 To UBound(m_binarr)
            PgmArr(LSBbit + i) = m_binarr(i)
            m_bitsum = m_bitsum + m_binarr(i)
        Next i
    Else
        ''''case:: LSBbit > MSBbit
        For i = 0 To UBound(m_binarr)
            PgmArr(LSBbit - i) = m_binarr(i)
            m_bitsum = m_bitsum + m_binarr(i)
        Next i
    End If

    auto_eFuse_HexStr2PgmArr_Write_byStage = m_HexStr
    
    ''''--------------------------------------------------------------------------------------------
    FuseType = UCase(Trim(FuseType))
    If (FuseType = "ECID") Then
        ''''Not Support,check it later
        ''TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (CFG,UID,UDR,SEN,MON)"
        ''GoTo errHandler
        
        ''''20170911 update and support
        ECIDFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        ECIDFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        ECIDFuse.Category(idx).Write.Decimal(ss) = m_decimal
        ECIDFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        ECIDFuse.Category(idx).Write.Value(ss) = m_decimal
        ECIDFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        ECIDFuse.Category(idx).Write.HexStr(ss) = m_HexStr

    ElseIf (FuseType = "CFG") Then
        CFGFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        CFGFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        CFGFuse.Category(idx).Write.Decimal(ss) = m_decimal
        CFGFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        CFGFuse.Category(idx).Write.Value(ss) = m_decimal
        CFGFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        CFGFuse.Category(idx).Write.HexStr(ss) = m_HexStr

    ElseIf (FuseType = "UID") Then
        UIDFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UIDFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UIDFuse.Category(idx).Write.Decimal(ss) = m_decimal
        UIDFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UIDFuse.Category(idx).Write.Value(ss) = m_decimal
        UIDFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UIDFuse.Category(idx).Write.HexStr(ss) = m_HexStr

    ElseIf (FuseType = "UDR") Then
        UDRFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UDRFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRFuse.Category(idx).Write.Decimal(ss) = m_decimal
        UDRFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UDRFuse.Category(idx).Write.Value(ss) = m_decimal
        UDRFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UDRFuse.Category(idx).Write.HexStr(ss) = m_HexStr

    ElseIf (FuseType = "SEN") Then
        SENFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        SENFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        SENFuse.Category(idx).Write.Decimal(ss) = m_decimal
        SENFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        SENFuse.Category(idx).Write.Value(ss) = m_decimal
        SENFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        SENFuse.Category(idx).Write.HexStr(ss) = m_HexStr

    ElseIf (FuseType = "MON") Then
        MONFuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        MONFuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        MONFuse.Category(idx).Write.Decimal(ss) = m_decimal
        MONFuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        MONFuse.Category(idx).Write.Value(ss) = m_decimal
        MONFuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        MONFuse.Category(idx).Write.HexStr(ss) = m_HexStr

    ElseIf (FuseType = "UDRE") Then
        UDRE_Fuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UDRE_Fuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRE_Fuse.Category(idx).Write.Decimal(ss) = m_decimal
        UDRE_Fuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UDRE_Fuse.Category(idx).Write.Value(ss) = m_decimal
        UDRE_Fuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UDRE_Fuse.Category(idx).Write.HexStr(ss) = m_HexStr

    ElseIf (FuseType = "UDRP") Then
        UDRP_Fuse.Category(idx).Write.BitStrM(ss) = m_bitStrM
        UDRP_Fuse.Category(idx).Write.BitStrL(ss) = StrReverse(m_bitStrM)
        UDRP_Fuse.Category(idx).Write.Decimal(ss) = m_decimal
        UDRP_Fuse.Category(idx).Write.BitSummation(ss) = m_bitsum
        UDRP_Fuse.Category(idx).Write.Value(ss) = m_decimal
        UDRP_Fuse.Category(idx).Write.ValStr(ss) = CStr(m_decimal)
        UDRP_Fuse.Category(idx).Write.HexStr(ss) = m_HexStr

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,UDRE,UDRP)"
        GoTo errHandler
    End If
    ''''--------------------------------------------------------------------------------------------

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20160519 Add
Public Function auto_BinStr2HexStr(ByVal BinStr As String, ByVal HexBit As Long) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_BinStr2HexStr"

    Dim i As Long, j As Long
    Dim BinStrLen As Long
    Dim HexMOD As Long
    Dim HexStr As String
    Dim HexVal As String
    Dim HexLen As Long

    HexStr = ""
    
    BinStrLen = Len(BinStr)
    If (BinStrLen Mod (4)) > 0 Then
        HexLen = (BinStrLen \ 4) + 1
    Else
        HexLen = BinStrLen \ 4
    End If
    
    If HexBit > HexLen Then
        HexLen = HexBit
    End If

    HexMOD = HexLen * 4 - BinStrLen
    
    If HexMOD > 0 Then
        For i = 0 To HexMOD - 1
            BinStr = "0" & BinStr
        Next i
    End If

    For i = 0 To HexLen - 1
        If Mid(BinStr, i * 4 + 1, 4) = "0000" Then
            HexVal = "0"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0001" Then
            HexVal = "1"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0010" Then
            HexVal = "2"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0011" Then
            HexVal = "3"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0100" Then
            HexVal = "4"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0101" Then
            HexVal = "5"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0110" Then
            HexVal = "6"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0111" Then
            HexVal = "7"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1000" Then
            HexVal = "8"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1001" Then
            HexVal = "9"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1010" Then
            HexVal = "A"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1011" Then
            HexVal = "B"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1100" Then
            HexVal = "C"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1101" Then
            HexVal = "D"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1110" Then
            HexVal = "E"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1111" Then
            HexVal = "F"
        Else
            HexVal = "X"
        End If

        HexStr = HexStr & HexVal
    Next i

    auto_BinStr2HexStr = HexStr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_eFuse_pwr_on_i_meter_DCVS(Pin As String, v As Double, i_rng As Double, _
                                               wait_before_gate As Double, wait_after_gate As Double, _
                                               Steps As Long, RiseTime As Double, _
                                               Optional showPrint As Boolean = False)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_pwr_on_i_meter_DCVS"

    'set Force voltage and Current/Meter Range
    ''===============================================
    ''Description: __                __
    ''                __|
    ''            __|
    ''        __|
    ''    __| >|   |<--stepT  __        v
    ''__|                   __ stepV   __
    ''|<-- steps -->
    ''|<-FallTime ->
    ''===============================================

    Dim i_meter_rng As Double
    i_meter_rng = i_rng
    
    Dim setV As Double
    Dim StepV As Double
    StepV = v / Steps

    Dim i As Long

    Dim stepT As Double
    stepT = RiseTime / Steps

    With TheHdw.DCVS.Pins(Pin)
        .Connect
        .mode = tlDCVSModeVoltage
        .Voltage.Main = 0
        .SetCurrentRanges i_rng, i_meter_rng
'        .CurrentLimit.Source.FoldLimit.Level = i_rng
        .Meter.mode = tlDCVSMeterCurrent
        .CurrentRange.Value = i_rng
        .CurrentLimit.Source.FoldLimit.Level.Value = i_rng
        .Meter.CurrentRange = i_rng
        TheHdw.Wait wait_before_gate   'wait for relay connect
        .Gate = True
    End With
    
    ''Pwr On Ramp up slew-rate control============================
    For i = 1 To Steps
        setV = i * StepV
        TheHdw.DCVS.Pins(Pin).Voltage.Main = setV
        
        If showPrint = True Then
            TheExec.Datalog.WriteComment "  Curr_" & Pin & " Pwr Up Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
        End If
        
        TheHdw.Wait stepT
    Next i
    ''============================================================

    TheHdw.Wait wait_after_gate

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_eFuse_pwr_off_i_meter_DCVS(Pin As String, v As Double, i_rng As Double, _
                                                wait_before_gate As Double, wait_after_gate As Double, _
                                                Steps As Long, FallTime As Double, _
                                                Optional showPrint As Boolean = False)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_pwr_off_i_meter_DCVS"

    'set Force voltage and Current/Meter Range
    ''===============================================
    ''Description
    ''__                              __
    ''  |__
    ''     |_>|  |<--stepT             v
    ''        |__          __
    ''           |__       __ stepV   __
    ''|<-- steps -->
    ''|<-FallTime ->
    ''===============================================

    Dim i_meter_rng As Double
    i_meter_rng = i_rng
    
    Dim setV As Double
    Dim StepV As Double
    StepV = v / Steps
    
    Dim i As Long
    Dim stepsm As Long
    
    Dim stepT As Double
    stepT = FallTime / Steps

    With TheHdw.DCVS.Pins(Pin)
        .Connect
        .mode = tlDCVSModeVoltage
        .Voltage.Main = v
        .SetCurrentRanges i_rng, i_meter_rng
'        .CurrentLimit.Source.FoldLimit.Level = i_rng
        .Meter.mode = tlDCVSMeterCurrent
        .CurrentRange.Value = i_rng
        .CurrentLimit.Source.FoldLimit.Level.Value = i_rng
        .Meter.CurrentRange = i_rng
        TheHdw.Wait wait_before_gate   'wait for relay connect
        .Gate = True
    End With
    
    ''Pwr On Ramp Down slew-rate control============================
    stepsm = Steps - 1
    For i = 0 To stepsm
        setV = v - (i * StepV)
        TheHdw.DCVS.Pins(Pin).Voltage.Main = setV
        
        If showPrint = True Then
            TheExec.Datalog.WriteComment "  Curr_" & Pin & " Pwr Down Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
        End If
        
        TheHdw.Wait stepT
    Next i

    setV = 0
    TheHdw.DCVS.Pins(Pin).Voltage.Main = setV
    
    If showPrint = True Then
        TheExec.Datalog.WriteComment "  Curr_" & Pin & " Pwr Down Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
    End If
    ''==============================================================

    TheHdw.Wait wait_after_gate

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160608 Add
''''20161215 update
Public Function CeilingValue(ByVal X As Double, Optional ByVal factor As Double = 1) As Variant

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CeilingValue"
    
    ' X is the value you want to round
    ' is the multiple to which you want to round up
    ''CeilingValue = (Int(X / Factor) - (X / Factor - Int(X / Factor) > 0)) * Factor
    Dim m_lngX As Long
    m_lngX = CLng(X)
    
    ''''20161215 update, By this way to avoid any unsuitable Ceiling value
    If ((CDbl(X) - m_lngX) > 0.000000001) Then
        CeilingValue = m_lngX + 1
    Else
        CeilingValue = m_lngX
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160617 New Function
Public Function auto_eFuse_copyLMTtoLMT_R() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_copyLMTtoLMT_R"

    Dim i As Long

    If (gB_findECID_flag) Then
        For i = 0 To UBound(ECIDFuse.Category)
            ECIDFuse.Category(i).HiLMT_R = ECIDFuse.Category(i).HiLMT
            ECIDFuse.Category(i).LoLMT_R = ECIDFuse.Category(i).LoLMT
        Next i
    End If

    If (gB_findUID_flag) Then
        For i = 0 To UBound(UIDFuse.Category)
            UIDFuse.Category(i).HiLMT_R = UIDFuse.Category(i).HiLMT
            UIDFuse.Category(i).LoLMT_R = UIDFuse.Category(i).LoLMT
        Next i
    End If

    If (gB_findCFG_flag) Then
        For i = 0 To UBound(CFGFuse.Category)
            CFGFuse.Category(i).HiLMT_R = CFGFuse.Category(i).HiLMT
            CFGFuse.Category(i).LoLMT_R = CFGFuse.Category(i).LoLMT
        Next i
    End If

    If (gB_findUDR_flag) Then
        For i = 0 To UBound(UDRFuse.Category)
            UDRFuse.Category(i).HiLMT_R = UDRFuse.Category(i).HiLMT
            UDRFuse.Category(i).LoLMT_R = UDRFuse.Category(i).LoLMT
        Next i
    End If

    If (gB_findSEN_flag) Then
        For i = 0 To UBound(SENFuse.Category)
            SENFuse.Category(i).HiLMT_R = SENFuse.Category(i).HiLMT
            SENFuse.Category(i).LoLMT_R = SENFuse.Category(i).LoLMT
        Next i
    End If

    If (gB_findMON_flag) Then
        For i = 0 To UBound(MONFuse.Category)
            MONFuse.Category(i).HiLMT_R = MONFuse.Category(i).HiLMT
            MONFuse.Category(i).LoLMT_R = MONFuse.Category(i).LoLMT
        Next i
    End If

    If (gB_findUDRE_flag) Then
        For i = 0 To UBound(UDRE_Fuse.Category)
            UDRE_Fuse.Category(i).HiLMT_R = UDRE_Fuse.Category(i).HiLMT
            UDRE_Fuse.Category(i).LoLMT_R = UDRE_Fuse.Category(i).LoLMT
        Next i
    End If

    If (gB_findUDRP_flag) Then
        For i = 0 To UBound(UDRP_Fuse.Category)
            UDRP_Fuse.Category(i).HiLMT_R = UDRP_Fuse.Category(i).HiLMT
            UDRP_Fuse.Category(i).LoLMT_R = UDRP_Fuse.Category(i).LoLMT
        Next i
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160620 New Function
''''20160926 MUST use ByVal for the variable vstr, otherwise the returned string will not have "0x".
Public Function auto_chkHexStr_isOver7FFFFFFF(ByVal vstr As String, Optional showPrint As Boolean = False) As Boolean

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_chkHexStr_isOver7FFFFFFF"

    Dim i As Long
    Dim m_len As Long
    Dim m_ch As String
    Dim m_1stch As String
    Dim m_1stchval As Long
    Dim m_result As Boolean
    
    vstr = UCase(vstr)
    vstr = Replace(vstr, "_", "") ''''20171103 update, for case like as 0x1234_ABCD_EFFE
    
    If (vstr Like "0X*") Then
        vstr = Replace(vstr, "0X", "", 1)

        ''''get rid of the zeros '0' before 1st Hex character
        ''''ex: 0x0002FF => vstr = "2FF"
        For i = 1 To Len(vstr)
            m_ch = Mid(vstr, 1, 1)
            If (m_ch = "0") Then
                vstr = Replace(vstr, "0", "", 1, 1)
            Else
                Exit For ''''<MUST>
            End If
        Next i
    ElseIf (vstr Like "X*") Then
        vstr = Replace(vstr, "X", "", 1)

        ''''get rid of the zeros '0' before 1st Hex character
        ''''ex: x0002FF => vstr = "2FF"
        For i = 1 To Len(vstr)
            m_ch = Mid(vstr, 1, 1)
            If (m_ch = "0") Then
                vstr = Replace(vstr, "0", "", 1, 1)
            Else
                Exit For ''''<MUST>
            End If
        Next i
    Else
        ''''<MUST>
        auto_chkHexStr_isOver7FFFFFFF = False
        Exit Function
    End If

    ''''20180522 update for allZero case
    If (vstr = "") Then vstr = "0"
    m_len = Len(vstr)

    If (m_len < 8) Then '''' <32bits
        m_result = False

    ElseIf (m_len = 8) Then '''' =32bits
        m_1stch = Mid(vstr, 1, 1)
        m_1stchval = CLng("&H" & m_1stch)
        If (m_1stchval > 7) Then
            m_result = True
        Else
            m_result = False
        End If
        
    Else '''' >32bits
        m_result = True
    End If

    auto_chkHexStr_isOver7FFFFFFF = m_result
    
    If (showPrint) Then
        TheExec.Datalog.WriteComment funcName + ":: is " + CStr(m_result) + " (" + CStr(vstr) + ")"
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160927, New
''''input hexStr MUST be with the prefix "0x".
''''20170811 update for case with the prefix "x".
''''20180522 update if HEX characters is over 255 (>1023bits) then keep the Hex as the output
Public Function auto_HexStr2Value(ByVal HexStr As String) As Variant

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_HexStr2Value"

    ''''-------------------------------------------------------------------------
    ''''Example, it could support up to (2^1024 -1)
    ''''-------------------------------------------------------------------------
    ''''Call auto_HexStr2Value("0x7FFFFFFF")
    ''''auto_HexStr2Value:: 0x7FFFFFFF = 2147483647
    ''''auto_HexStr2Value:: 0xFFFF = 65535
    ''''auto_HexStr2Value:: 0x2F = 47
    ''''Call auto_HexStr2Value("0x80000000")
    ''''auto_HexStr2Value:: 0x80000000 = 2147483648
    ''''Call auto_HexStr2Value("0xFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
    ''''auto_HexStr2Value::     0xFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF = 1.34078079299426E+154
    ''''Call auto_HexStr2Value("0xFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
    ''''auto_HexStr2Value::     0xFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF = 1.15792089237316E+77
    ''''auto_HexStr2Value::     0xFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF = 3.40282366920938E+38
    ''''auto_HexStr2Value::     0xFFFFFFFFFFFFFFFF = 1.84467440737096E+19
    ''''auto_HexStr2Value::     0x8FFFFFFF = 2415919103
    ''''-------------------------------------------------------------------------

    Dim i As Long
    Dim m_HexStr As String ''''without the prefix '0x' or 'x'
    Dim m_len As Long
    Dim m_char As String
    Dim m_chVal As Double
    Dim m_hex2Val As Double
    
    HexStr = UCase(Trim(HexStr))
    m_hex2Val = 0
    HexStr = Replace(HexStr, "_", "") ''''20171103 update, for case like as 0x1234_ABCD_EFFE

    If (auto_isHexString(HexStr) = True) Then
        ''''20170811 update
        If (HexStr Like "0X*") Then
            m_HexStr = Replace(HexStr, "0X", "", 1)
        ElseIf (HexStr Like "X*") Then
            m_HexStr = Replace(HexStr, "X", "", 1)
        End If
        m_len = Len(m_HexStr)
        For i = 1 To m_len
            m_char = Mid(m_HexStr, i, 1)
            Select Case m_char
                Case "A"
                    m_chVal = 10
                Case "B"
                    m_chVal = 11
                Case "C"
                    m_chVal = 12
                Case "D"
                    m_chVal = 13
                Case "E"
                    m_chVal = 14
                Case "F"
                    m_chVal = 15
                Case Else
                    m_chVal = CDbl(m_char)
            End Select
            ''''20180522 update if HEX characters is over 255 (>1023bits)
            If ((m_len - i) < 256) Then
                m_hex2Val = m_hex2Val + m_chVal * (16 ^ (m_len - i))
            ElseIf ((m_len - i) >= 256 And m_chVal = 0) Then
                m_hex2Val = m_hex2Val + 0
            Else
                auto_HexStr2Value = HexStr
                ''TheExec.AddOutput "<WARNING> " + funcName + ":: " + HexStr + " is over 1023bits(255 Hex_Characters)."
                ''TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: " + HexStr + " is over 1023bits(255 Hex_Characters)."
                ''Debug.Print funcName + ":: " + HexStr + " = " + CStr(auto_HexStr2Value)
                Exit Function
            End If
        Next i
        auto_HexStr2Value = m_hex2Val
        ''TheExec.Datalog.WriteComment funcName + ":: " + hexStr + " = " + CStr(m_hex2Val)
        ''Debug.Print funcName + ":: " + HexStr + " = " + CStr(m_hex2Val)
    Else
        auto_HexStr2Value = 0
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: " + HexStr + " is NOT a Hex String (with the prefix '0x' or 'x')."
        GoTo errHandler
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20161114 update for validate/load the pattern in the OnProgramValidation
''''showTime = True/False to display the test time
Public Function auto_eFuse_PatSetToPat_Validation(ByVal patset As Pattern, ByRef patt As String, Optional Validating_ As Boolean, _
Optional showTime As Boolean = False) As Boolean

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_PatSetToPat_Validation"

    If (showTime) Then Call auto_StartWatchTimer
    Dim i As Long
    Dim PatAry() As String, PatCnt As Long
    Dim m_patset As New Pattern
    Dim m_pat As String
    
    ''''20171211 add, to prevent Unexpected pattern DSSC with Empty data
    TheHdw.Digital.Patgen.Halt ''''<MUST and Could be>

    ''''------------------------------------------------------------------------
    ''''Parsing PatternSet to get the raw pattern name (.pat)
    ''''Actually, eFuse PatternSet only contains one pat file.
    ''''------------------------------------------------------------------------
    'PatAry = TheExec.DataManager.Raw.GetPatternsInSet(PatSet, patCnt)
    If (LCase(patset.Value) Like "*.pat") Then
        ReDim PatAry(0)
        PatAry(0) = patset.Value
    Else
        PatAry = TheExec.DataManager.Raw.GetPatternsInSet(patset, PatCnt)
    End If
    
    While Not (LCase(PatAry(0)) Like "*.pat*")
        m_patset.Value = PatAry(0)
        PatAry = TheExec.DataManager.Raw.GetPatternsInSet(m_patset, PatCnt)
        If UBound(PatAry) > 1 Then TheExec.ErrorLogMessage (patset & " is with more than one pattern in the pattern set")
    Wend
    patt = PatAry(0)
    ''''<NOTICE> lowcase ".gz" will be implicit, so pattern(.gz).load/test will be problem.
    ''''20161124 update to prevent *.gz (lowcase)
    If (patt Like "*.gz") Then
        patt = Replace(patt, ".gz", "")
    End If
    ''''------------------------------------------------------------------------
    
    auto_eFuse_PatSetToPat_Validation = False ''''init
    If (Validating_) Then
        ''''<MUST> By this way, the PatSet can be explicit in the pattern memory.
        If (ValidatePattern(patset.Value) = False) Then
            TheExec.AddOutput "<Error> " + funcName + ":: please check the PatternSet, " + CStr(patset.Value)
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check the PatternSet, " + CStr(patset.Value)
            GoTo errHandler
        End If

        ''''Actually, here it's just only one pattern.
        ''''<MUST> By this way, the individual pattern (.pat) can be explicit in the pattern memory.
        For i = 0 To UBound(PatAry)
            m_pat = PatAry(i)
            If (ValidatePattern(m_pat) = False) Then
                TheExec.AddOutput "<Error> " + funcName + ":: please check the Pattern, " + CStr(m_pat)
                TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check the Pattern, " + CStr(m_pat)
                GoTo errHandler
            End If
        Next i
        auto_eFuse_PatSetToPat_Validation = True ''''<MUST>
        If (showTime) Then Call auto_StopWatchTimer(funcName)
        Exit Function ''''<MUST>
    End If

    For i = 0 To UBound(PatAry)
        m_pat = PatAry(i)
        ''''<NOTICE> lowcase ".gz" will be implicit, so pattern(.gz).load/test will be problem.
        ''''20161124 update to prevent *.gz (lowcase)
        If (m_pat Like "*.gz") Then
            m_pat = Replace(m_pat, ".gz", "")
        End If
        
        '''<MUST> put it here when user unloadAllPatterns, it can reload the pattern.
        '''<MUST> DSSC Src/Cap setup needs this statement
        TheHdw.Patterns(m_pat).Load ''''it will take 0.3-0.5 ms if the pattern was loaded
    Next i

    TheHdw.Wait 0.00001 ''''10uS
    If (showTime) Then Call auto_StopWatchTimer(funcName)

Exit Function
errHandler:
    TheExec.AddOutput "<Error> " + funcName + ":: please check it out."
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20161114 update
Public Function auto_StartWatchTimer()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_StartWatchTimer"

    TheHdw.StartStopwatch

Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

''''20161114 update
Public Function auto_StopWatchTimer(Optional itemStr As String = "", Optional unit_ms As Boolean = True)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_StopWatchTimer"
    
    Dim m_instName As String
    Dim m_tmpStr As String
    Dim m_timedbl As Double
    Dim m_len As Long

    m_timedbl = TheHdw.ReadStopwatch
    
    m_instName = FormatNumeric(TheExec.DataManager.instanceName, 35)
    itemStr = FormatNumeric(itemStr, 10)

    If (Trim(m_instName) <> "" And Trim(itemStr) <> "") Then
        m_tmpStr = m_instName + "::" + itemStr + " = "
    ElseIf (Trim(m_instName) = "" And Trim(itemStr) <> "") Then
        m_tmpStr = itemStr + " = "
    ElseIf (Trim(m_instName) <> "" And Trim(itemStr) = "") Then
        m_tmpStr = m_instName + " = "
    Else
        m_tmpStr = ""
    End If
    m_len = Len(m_tmpStr)
    
    If (unit_ms) Then
        m_tmpStr = vbTab & "Test Time " + FormatNumeric(m_tmpStr, m_len) + Format(m_timedbl * 1000, "0.0000") + " mS."
    Else
        m_tmpStr = vbTab & "Test Time " + FormatNumeric(m_tmpStr, m_len) + Format(m_timedbl, "0.000000") + " Secs."
    End If
    ''TheExec.Datalog.WriteComment m_tmpStr
    Debug.Print m_tmpStr
    ''''If (UCase(m_tmpStr) Like UCase("*Validation*")) Then Debug.Print m_tmpStr
    
    ''''using for the next/continue stopWatch to avoid the above code statements (extral time)
    TheHdw.StartStopwatch
    
Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

''''20161114 copy from Dec2BinStr32Bit in LIB_Common()
Public Function auto_eFuse_Dec2BinStr32Bit(ByVal Nbit As Long, ByVal num As Long) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Dec2BinStr32Bit"

    ' 2'complement: invert the number's bits and then add 1
    'Dec2BinStr32Bit 32, -65525
    '1111111111111110000000000001011    -65525
    '0000000000000001111111111110101     65525
    Dim i As Long, j As Long
    Dim Element_Amount As Long
    Dim Count As Long
    Dim BinStr As String
    ' MSB "010101" LSB
    
    BinStr = ""
    If Nbit < 1 Then MsgBox ("Warning(" + funcName + ")!!! Decimal Number or number of Bit is wrong")
    If Nbit = 32 Then
        Nbit = 30
        If num < 0 Then
            BinStr = "1"
        Else
            BinStr = "0"
        End If
    End If
    For i = Nbit To 0 Step -1
        If num And (2 ^ i) Then
            BinStr = BinStr & "1"
        Else
            BinStr = BinStr & "0"
        End If
    Next
    auto_eFuse_Dec2BinStr32Bit = BinStr
    ''''Debug.Print BinStr

Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

''''20161121, copy from DSSC_DigCapSetup in module LIB_Common_DSSC_HRAM
''''remove the redundant pattern.load and eFuse usage only
Public Function auto_eFuse_DSSC_DigCapSetup(patt As String, DigCapPin As PinList, SignalName As String, SampleSize As Long, DspWav As DSPWave)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_DSSC_DigCapSetup"

    ''''TheHdw.Patterns(Patt).Load
    With TheHdw.DSSC.Pins(DigCapPin).Pattern(patt).Capture.Signals
        .Add (SignalName)
        With .Item(SignalName)
            .SampleSize = SampleSize 'CaptureCyc * OneCycle
            .LoadSettings
        End With
    End With
    
    'capture
    DspWav = TheHdw.DSSC.Pins(DigCapPin).Pattern(patt).Capture.Signals(SignalName).DSPWave
    
    'Bypass DSP computing, use HOST computer
    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug ''tlDSPModeHostDebug,tlDSPModeAutomatic
    'halt on opcode to make sure all samples are capture.
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode

Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

'''''20161114, copy from DSSC_ReadDigCap in module LIB_Common_DSSC_HRAM
'''''It's used in 32bits read of eFuse
'''was Public Function auto_eFuse_DSSC_ReadDigCap_32bits(CycleNum As Long, PinName As String, ByRef SingleStrArray() As String, CapWave As DSPWave, ByRef blank As SiteBoolean)
'''''20170126 update the feature, serial capture then transfer to 32bits binary string
'''''20180122 update from Tosp
'Public Function auto_eFuse_DSSC_ReadDigCap_32bits(cycleNum As Long, PinName As String, ByRef SingleStrArray() As String, CapWave As DSPWave, ByRef blank As SiteBoolean, _
'                                                  Optional serialCap As Boolean = False, Optional PatBitOrder As String = "bit0_bitLast")
'    If FunctionList.Exists("auto_eFuse_DSSC_ReadDigCap_32bits") = False Then FunctionList.Add "auto_eFuse_DSSC_ReadDigCap_32bits", ""
'
'On Error GoTo errHandler
'    Dim funcName As String:: funcName = "auto_eFuse_DSSC_ReadDigCap_32bits"
'
'    Dim Pin_Ary() As String, Pin_Cnt As Long, p_indx As Long, N1 As Long
'    Dim ByteString As String, CurrInstance As String
'    Dim hram_pindata As New PinListData
'    Dim Cdata As String
'    Dim p As Variant, p_idx As Long
'    Dim AlarmStr As String
'    Dim Site As Variant
'    ReDim SingleStrArray(cycleNum - 1, TheExec.Sites.Existing.Count - 1)
'
'    'initial
'    blank = True
'
'    If (serialCap = False) Then
'        For Each Site In TheExec.Sites
'            For N1 = 0 To cycleNum - 1
'                ByteString = auto_eFuse_Dec2BinStr32Bit(32, CapWave.Element(N1))
'                If CapWave.Element(N1) <> 0 Then blank(Site) = False    'return blank bolean for ecid blank check
'                SingleStrArray(N1, Site) = ByteString
'                ''Debug.Print N1 & ", " & CapWave.Element(N1) & " = ", ByteString
'            Next N1
'        Next Site
'    Else
'        ''''20180122 update to support Serial to Parallel
'        ''''Serial Capture and transfer to 32bits Binary String
'        Call auto_eFuse_DSSC_ReadDigCap_1bit_to_32bits(cycleNum, SingleStrArray, CapWave, blank, PatBitOrder)
'    End If
'    TheHdw.Wait 0.001
'
'Exit Function
'errHandler:
'     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'     If AbortTest Then Exit Function Else Resume Next
'End Function
'
''''20161114 update
Public Function auto_eFuse_to_STDF_allBits(ByVal FuseType As String, binStrM As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_to_STDF_allBits"

    Dim ss As Variant
    Dim m_THStr As String
    Dim m_TSStr As String
    Dim m_HexStr As String
    Dim m_hexlen As Long
    Dim m_len As Long
    Dim m_tmpStr As String
    
    Dim i As Long
    Dim m_lenFactor As Long
    Dim m_startbit As Long
    Dim m_stopbit As Long
    Dim m_hexstrP As String ''''partial hexstr
    Dim m_block As Long

    ss = TheExec.sites.SiteNumber

    ''''----------------------------------------------------
    ''''eFuse ECID/Config/UDR/MON be placed in a custom DTR with following fields,
    '''' FUSE,fusetype,TH,TS,MIN,MAX,<DATA>
    ''''- FUSE is just the naming.
    ''''- ECID/Config/UDR/SEN/MON to specify which one.
    ''''- TH is test head
    ''''- TS is test site
    ''''- MIN/MAX the bit start/stop
    ''''- The data. (Hex)
    ''''---------------------------------------------------

    ''''<NOTICE> it's needed to check it later.
    ''m_THStr = Environ$("computername")
    m_THStr = "1"
    m_TSStr = CStr(ss)

    FuseType = UCase(Trim(FuseType))
    If (FuseType Like "*CFG*") Then
        FuseType = "CONFIG"
    End If

    m_len = Len(binStrM)
    m_hexlen = IIf(m_len - (m_len \ 4) * 4 = 0, (m_len \ 4), (m_len \ 4) + 1)
    m_HexStr = auto_BinStr2HexStr(binStrM, m_hexlen)

    ''TheExec.Datalog.WriteComment ""
    ''''20161206 update to avoid the long hex string to be truncated if the string length > m_lenFactor.
    m_lenFactor = 128
    If (m_hexlen <= m_lenFactor) Then
        m_tmpStr = "FUSE," + FuseType + "," + m_THStr + "," + m_TSStr + ",0," + CStr(m_len - 1) + "," + m_HexStr
        TheExec.Datalog.WriteComment m_tmpStr ''''it's DTR in STDF, HexStr[MSB......LSB]
    Else
        ''''over 128 Hex characters (=128x4=512 bits)
        m_block = Ceiling(m_hexlen / m_lenFactor) ''''mean how many hex block (base=m_lenFactor)
        For i = 0 To (m_block - 1)
            If (i = (m_block - 1)) Then ''''last remainder hexstr
                m_startbit = i * m_lenFactor * 4
                m_stopbit = m_startbit + (m_hexlen - (i * m_lenFactor)) * 4 - 1
                m_hexstrP = StrReverse(Mid(StrReverse(m_HexStr), 1 + (i * m_lenFactor), (m_hexlen - (i * m_lenFactor))))
            Else
                m_startbit = i * m_lenFactor * 4
                m_stopbit = m_startbit + (m_lenFactor * 4) - 1
                m_hexstrP = StrReverse(Mid(StrReverse(m_HexStr), 1 + (i * m_lenFactor), m_lenFactor))
            End If
            m_tmpStr = "FUSE," + FuseType + "," + m_THStr + "," + m_TSStr + "," + CStr(m_startbit) + "," + CStr(m_stopbit) + "," + m_hexstrP
            TheExec.Datalog.WriteComment m_tmpStr ''''it's DTR in STDF, HexStr[MSB......LSB]
        Next i
    End If

    TheExec.Datalog.WriteComment ""

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

''''20161205 update
Public Function auto_Compare_EQN_Voltage_Per_Site(ids_current As Double, p_mode As Integer, GRADEVDD As Double, resolution As Double, PassBin As Long, EQN_Number As Long)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Compare_EQN_Voltage_Per_Site"

    ''Dim Site As Variant
    Dim passbincut_num As Variant
    Dim step As Long
    Dim p As Long, i As Long, k As Long
    Dim remainder As Double
    Dim divisor As Double
    Dim cal_voltage As Double
    Dim final_voltage As Double
    Dim showPrint As Boolean
    
    '////////////////////////////////////////
    divisor = resolution ''''gC_StepVoltage
    '////////////////////////////////////////

    showPrint = True ''''for Debug
    
    If BinCut(p_mode, PassBin).ExcludedPmode = False Then
        If (showPrint) Then TheExec.Datalog.WriteComment "Total Mode_Step=" & (BinCut(p_mode, PassBin).Mode_Step + 1)
        ''''As is : [BinCut(P_mode, PASSBIN).Mode_Step - 1], it will cause that the search is failure.
        For k = 0 To BinCut(p_mode, PassBin).Mode_Step
            final_voltage = 0
            cal_voltage = BinCut(p_mode, PassBin).c(k) - BinCut(p_mode, PassBin).M(k) * (Log(ids_current * 1000) / Log(10))
            ''''debug print
            'If (showPrint) Then TheExec.Datalog.WriteComment "cal_voltage=" & cal_voltage
            remainder = cal_voltage / divisor
            'If (showPrint) Then TheExec.Datalog.WriteComment "Remainder=" & Remainder
            remainder = Floor(remainder)
            'If (showPrint) Then TheExec.Datalog.WriteComment "Floor(Remainder)=" & Remainder
            cal_voltage = remainder * divisor
            'If (showPrint) Then TheExec.Datalog.WriteComment "final cal_voltage=" & cal_voltage
            If cal_voltage > BinCut(p_mode, PassBin).CP_Vmax(k) Then
                  final_voltage = BinCut(p_mode, PassBin).CP_Vmax(k) + BinCut(p_mode, PassBin).CP_GB(k)
            Else
                  If cal_voltage < BinCut(p_mode, PassBin).CP_Vmin(k) Then
                      final_voltage = BinCut(p_mode, PassBin).CP_Vmin(k) + BinCut(p_mode, PassBin).CP_GB(k)
                  Else
                      final_voltage = cal_voltage + BinCut(p_mode, PassBin).CP_GB(k)
                  End If
            End If
            ''''debug print
            If (showPrint) Then
                TheExec.Datalog.WriteComment "EQ_" & (k + 1) & ", GRADEVDD=" & GRADEVDD & ", IDS_current=" & ids_current & ", Final_voltage=" & final_voltage
            End If
            If GRADEVDD = final_voltage Then
                EQN_Number = k + 1
                Exit For
            End If
        Next k
    Else
        EQN_Number = 999 ''''set to failure
        TheExec.ErrorLogMessage "The Performance Power for " & VddBinName(p_mode) & " does not exist"
    End If

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''---------------------------------------------------------------------------
''''20170630 New CFG_Condition_Table structure for Tcyp and later.
''''201811XX Update with C00~C12 and D00~D12 case
''''         and support non-continuous condition bits as MC2T project
''''---------------------------------------------------------------------------
Public Function parse_CFG_Condition_Table_Sheet(sheetName As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "parse_CFG_Condition_Table_Sheet"
    
    Const init_arrSize = 1024
    Dim mysheet As Worksheet
    Dim myCell As Object
    Dim offCell As Object
    Dim myCell_Header As Object
    
    Dim myCellA1 As Object
    Dim m_A1_rowCnt As Long
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim M As Long
    Dim n As Long
    Dim m_cellCnt As Long
    Dim m_cellStr As String
    Dim m_offcolStr As String
    Dim m_lastrow As Long
    Dim m_lastNCnt As Long

    Dim find_1stHeader As Boolean
    Dim find_AllHeader As Boolean
    
    Dim idx_1stHeader_row As Long
    Dim idx_Condition As Long
    Dim idx_MSB As Long
    Dim idx_LSB As Long
    Dim idx_BitWidth As Long
    Dim idx_Stage As Long
    
    Dim idx_Default As Long
    Dim idx_ALL_0 As Long  ''''all elements are Zero
    Dim idx_1st_SI As Long ''''1st Silicon
    Dim idx_A00 As Long
    Dim idx_A01 As Long
    Dim idx_A02 As Long
    Dim idx_A03 As Long
    Dim idx_A04 As Long
    Dim idx_A05 As Long
    Dim idx_A06 As Long
    Dim idx_A07 As Long
    Dim idx_A09 As Long
    Dim idx_A12 As Long
    Dim idx_CommentA As Long

    Dim idx_B00 As Long
    Dim idx_B01 As Long
    Dim idx_B02 As Long
    Dim idx_B03 As Long
    Dim idx_B04 As Long
    Dim idx_B05 As Long
    Dim idx_B06 As Long
    Dim idx_B07 As Long
    Dim idx_B09 As Long
    Dim idx_B12 As Long
    Dim idx_CommentB As Long

    ''''201811XX add CFG_CXX and CFG_DXX
    Dim idx_C00 As Long
    Dim idx_C01 As Long
    Dim idx_C02 As Long
    Dim idx_C03 As Long
    Dim idx_C04 As Long
    Dim idx_C05 As Long
    Dim idx_C06 As Long
    Dim idx_C07 As Long
    Dim idx_C09 As Long
    Dim idx_C12 As Long
    Dim idx_CommentC As Long
    
    Dim idx_D00 As Long
    Dim idx_D01 As Long
    Dim idx_D02 As Long
    Dim idx_D03 As Long
    Dim idx_D04 As Long
    Dim idx_D05 As Long
    Dim idx_D06 As Long
    Dim idx_D07 As Long
    Dim idx_D09 As Long
    Dim idx_D12 As Long
    Dim idx_CommentD As Long
    Dim idx_END As Long
    
    Dim m_rowNum As Long
    Dim m_colNum As Long
    Dim m_cateCnt As Long
    Dim m_tmpStr As String
    Dim m_MSBBit As Long
    Dim m_LSBbit As Long
    Dim m_bitwidth As Long
    ''''-------------------------------
    ''''Debug Purpose
    ''''-------------------------------
    DebugPrtImm = False
    DebugPrtDlog = False
    ''''-------------------------------

    gB_findCFGCondTable_flag = False
    
    Set mysheet = Sheets(sheetName)

    Set myCellA1 = mysheet.range("A1")
    m_A1_rowCnt = 0
    
    Set myCell = mysheet.range("A1")
    m_cellStr = UCase(Trim(myCell.Value))
    
    DebugPrintLog "Row=" & myCell.row & ", Column=" & myCellA1.Column & ", End_Column=" & myCellA1.End(xlToRight).Column
    DebugPrintLog "A1...Row=" & myCell.row & ", Column=" & myCell.Column & ", Cell=" & myCell.Value & " (m_cellStr=" + m_cellStr + ")"
    
    find_1stHeader = False
    find_AllHeader = False
    m_A1_rowCnt = 0

    ''''At First, finding all Header Index to identify every column.
    ''''Here 'Do...Loop While()' is to search from Up to Down
    Do
        DebugPrintLog "0...Row=" & myCell.row & ", Column=" & myCell.Column & ", Cell=" & myCell.Value & " (m_cellStr=" + m_cellStr + ")"

        ''''1......To find the 1st Word "Condition"
        ''''Here Search Cell from Left to Right
        If (find_1stHeader = False) Then
            m_cellCnt = 0
            Do While (m_cellCnt < 11)
                ''DebugPrintLog "1...Row=" & myCell.Row & ", Column=" & myCell.Column & ", Cell=" & myCell.Value & " (m_cellStr=" + m_cellStr + ")"
                If (m_cellStr = UCase("Condition")) Then
                    idx_1stHeader_row = myCell.row
                    idx_Condition = myCell.Column
                    
                    DebugPrintLog "1...(Category) Row=" & idx_1stHeader_row & ", Column=" & idx_Condition & ", Cell=" & myCell.Value & " (m_cellStr=" + m_cellStr + ")"
                    find_1stHeader = True
                    Exit Do
                End If
    
                ''''if cell search from left to right, (rowOffset:=0, columnOffset:=1)
                Set myCell = myCell.Offset(rowOffset:=0, columnOffset:=1)
                m_cellStr = UCase(Trim(myCell.Value))
                m_cellCnt = m_cellCnt + 1
            Loop
        End If

        ''''2......To find the following Header Words
        ''''By each Header, get the related parameters.
        If (find_1stHeader) Then
            ''''<MUST> Point to the 1st cell of the row which contains the 1st Header Word
            Set myCell = myCellA1.Offset(rowOffset:=idx_1stHeader_row - 1, columnOffset:=0)
            m_cellStr = UCase(Trim(myCell.Value))
            DebugPrintLog "2...(find_1stHeader=True) Row=" & myCell.row & ", Column=" & myCell.Column & ", Cell=" & myCell.Value & " (m_cellStr=" + m_cellStr + ")"

            M = 0
            
            Do While (m_offcolStr <> "END")
                ''''if cell search from left to right, (rowOffset:=0, columnOffset:=1)
                Set offCell = myCell.Offset(rowOffset:=0, columnOffset:=M) ''''search cell from Left to Right
                m_offcolStr = UCase(Trim(offCell.Value))
                M = M + 1
                ''''-------------------------
                ''''Column Sequence
                ''''-------------------------
                ''''Condition
                ''''MSB Bit
                ''''LSB Bit
                ''''Bit Width
                ''''programming stage
                ''''A00
                ''''A01
                ''''...
                ''''A12
                ''''CommentA
                ''''B00
                ''''B01
                ''''...
                ''''B12
                ''''CommentB
                ''''C00
                ''''C01
                ''''...
                ''''C12
                ''''CommentC
                ''''D00
                ''''D01
                ''''...
                ''''D12
                ''''CommentD
                ''''End
                ''''-------------------------

                If (m_offcolStr = UCase("Condition")) Then
                    idx_Condition = M
                ElseIf (m_offcolStr Like UCase("MSB*Bit")) Then
                    idx_MSB = M
                ElseIf (m_offcolStr Like UCase("LSB*Bit")) Then
                    idx_LSB = M
                ElseIf (m_offcolStr Like UCase("Bit*Width")) Then
                    idx_BitWidth = M
                ElseIf (m_offcolStr Like UCase("*stage*")) Then
                    idx_Stage = M
                ''''-----------------------------------------------
                ElseIf (m_offcolStr = UCase("DEFAULT")) Then
                    idx_Default = M
                ElseIf (m_offcolStr = UCase("ALL_0")) Then
                    idx_ALL_0 = M
                ElseIf (m_offcolStr = UCase("1st_SI")) Then
                    idx_1st_SI = M
                ''''-----------------------------------------------
                ElseIf (m_offcolStr = UCase("A00")) Then
                    idx_A00 = M
                ElseIf (m_offcolStr = UCase("A01")) Then
                    idx_A01 = M
                ElseIf (m_offcolStr = UCase("A02")) Then
                    idx_A02 = M
                ElseIf (m_offcolStr = UCase("A03")) Then
                    idx_A03 = M
                ElseIf (m_offcolStr = UCase("A04")) Then
                    idx_A04 = M
                ElseIf (m_offcolStr = UCase("A05")) Then
                    idx_A05 = M
                ElseIf (m_offcolStr = UCase("A06")) Then
                    idx_A06 = M
                ElseIf (m_offcolStr = UCase("A07")) Then
                    idx_A07 = M
                ElseIf (m_offcolStr = UCase("A09")) Then
                    idx_A09 = M
                ElseIf (m_offcolStr = UCase("A12")) Then
                    idx_A12 = M
                ElseIf (m_offcolStr = UCase("CommentA")) Then
                    idx_CommentA = M
                ''''-----------------------------------------------
                ElseIf (m_offcolStr = UCase("B00")) Then
                    idx_B00 = M
                ElseIf (m_offcolStr = UCase("B01")) Then
                    idx_B01 = M
                ElseIf (m_offcolStr = UCase("B02")) Then
                    idx_B02 = M
                ElseIf (m_offcolStr = UCase("B03")) Then
                    idx_B03 = M
                ElseIf (m_offcolStr = UCase("B04")) Then
                    idx_B04 = M
                ElseIf (m_offcolStr = UCase("B05")) Then
                    idx_B05 = M
                ElseIf (m_offcolStr = UCase("B06")) Then
                    idx_B06 = M
                ElseIf (m_offcolStr = UCase("B07")) Then
                    idx_B07 = M
                ElseIf (m_offcolStr = UCase("B09")) Then
                    idx_B09 = M
                ElseIf (m_offcolStr = UCase("B12")) Then
                    idx_B12 = M
                ElseIf (m_offcolStr = UCase("CommentB")) Then
                    idx_CommentB = M
                ''''-----------------------------------------------
                ElseIf (m_offcolStr = UCase("C00")) Then
                    idx_C00 = M
                ElseIf (m_offcolStr = UCase("C01")) Then
                    idx_C01 = M
                ElseIf (m_offcolStr = UCase("C02")) Then
                    idx_C02 = M
                ElseIf (m_offcolStr = UCase("C03")) Then
                    idx_C03 = M
                ElseIf (m_offcolStr = UCase("C04")) Then
                    idx_C04 = M
                ElseIf (m_offcolStr = UCase("C05")) Then
                    idx_C05 = M
                ElseIf (m_offcolStr = UCase("C06")) Then
                    idx_C06 = M
                ElseIf (m_offcolStr = UCase("C07")) Then
                    idx_C07 = M
                ElseIf (m_offcolStr = UCase("C09")) Then
                    idx_C09 = M
                ElseIf (m_offcolStr = UCase("C12")) Then
                    idx_C12 = M
                ElseIf (m_offcolStr = UCase("CommentC")) Then
                    idx_CommentC = M
                ''''-----------------------------------------------
                ElseIf (m_offcolStr = UCase("D00")) Then
                    idx_D00 = M
                ElseIf (m_offcolStr = UCase("D01")) Then
                    idx_D01 = M
                ElseIf (m_offcolStr = UCase("D02")) Then
                    idx_D02 = M
                ElseIf (m_offcolStr = UCase("D03")) Then
                    idx_D03 = M
                ElseIf (m_offcolStr = UCase("D04")) Then
                    idx_D04 = M
                ElseIf (m_offcolStr = UCase("D05")) Then
                    idx_D05 = M
                ElseIf (m_offcolStr = UCase("D06")) Then
                    idx_D06 = M
                ElseIf (m_offcolStr = UCase("D07")) Then
                    idx_D07 = M
                ElseIf (m_offcolStr = UCase("D09")) Then
                    idx_D09 = M
                ElseIf (m_offcolStr = UCase("D12")) Then
                    idx_D12 = M
                ElseIf (m_offcolStr = UCase("CommentD")) Then
                    idx_CommentD = M
                ''''-----------------------------------------------
                ElseIf (m_offcolStr = UCase("END")) Then
                    idx_END = M
                End If
                
                DebugPrintLog "3...input m_offcolStr=" + m_offcolStr + " (" & offCell.Value & "), Index(Column)=" + CStr(M)
                
            Loop ''''end of Do While (m_offcolStr <> "END")
            
            If (m_offcolStr = ("END")) Then
                idx_END = M
                find_AllHeader = True
            End If
        End If

        ''''if cell search from up   to down,  (rowOffset:=1, columnOffset:=0)
        ''''if cell search from left to right, (rowOffset:=0, columnOffset:=1)
        ''''Here it MUST be use A1 as reference cell
        m_A1_rowCnt = m_A1_rowCnt + 1
        Set myCell = myCellA1.Offset(rowOffset:=m_A1_rowCnt, columnOffset:=0)
        m_cellStr = UCase(Trim(myCell.Value))
    Loop While (find_AllHeader = False)
    
    ''''After getting all Headers, process/get every content inside.
    If (find_AllHeader) Then
    
        ''''<MUST> Point to the 1st cell of the row which contains the 1st Header Word
        Set myCell_Header = myCellA1.Offset(rowOffset:=idx_1stHeader_row - 1, columnOffset:=0)
        m_cellStr = UCase(Trim(myCell_Header.Value))
        DebugPrintLog "4...(find_AllHeader=True) Row=" & myCell_Header.row & ", Column=" & myCell_Header.Column & ", Cell=" & myCell_Header.Value & " (m_cellStr=" + m_cellStr + ")"
        
        ''''initialize -----------------------------------------
        ReDim CFGTable.Category(init_arrSize)
        ''''----------------------------------------------------
        
        ''''-----------------------------------------------------------------------------------------------------
        ''''Get all PKG Names
        ''''-----------------------------------------------------------------------------------------------------
        M = 0
        m_cateCnt = 0
        Set myCell = myCell_Header.Offset(rowOffset:=0, columnOffset:=0) ''''rowOffset MUST be always '0'
        m_cellStr = UCase(Trim(myCell.Value))
        m_lastrow = myCell.End(xlDown).row
        m_lastNCnt = m_lastrow - idx_1stHeader_row
        DebugPrintLog "5...input Header =" + m_cellStr + " (" & myCell.Value & "), m=" + CStr(M) + ", LastRow=" + CStr(m_lastrow) + ", LastNCnt=" + CStr(m_lastNCnt)

        ''''case like "DEFAULT"/"ALL_0"/"1ST_SI"/"A##"/"B##"/"C##"/"D##" (#:0...9)
        Do
            If (m_cellStr = "DEFAULT" Or m_cellStr Like "ALL_0" Or m_cellStr Like "1ST_SI" Or _
                m_cellStr Like "A##" Or m_cellStr Like "B##" Or m_cellStr Like "C##" Or m_cellStr Like "D##") Then
                m_rowNum = myCell.row
                m_colNum = myCell.Column
                CFGTable.Category(m_cateCnt).pkgName = m_cellStr
                CFGTable.Category(m_cateCnt).FuseName = m_cellStr
                CFGTable.Category(m_cateCnt).row = m_rowNum
                CFGTable.Category(m_cateCnt).col = m_colNum
                DebugPrintLog "5-1...CFGTable.Category(" & m_cateCnt & ").PKGName=" & m_cellStr & ", row=" & m_rowNum & ", column=" & m_colNum
                m_cateCnt = m_cateCnt + 1
            End If
            Set myCell = myCell.Offset(rowOffset:=0, columnOffset:=1) ''''search cell from Left to Right
            m_cellStr = UCase(Trim(myCell.Value))
        Loop While (m_cellStr <> "END")
        ReDim Preserve CFGTable.Category(m_cateCnt - 1) ''''<MUST>
        ''''-----------------------------------------------------------------------------------------------------
        
        
        ''''initialize -----------------------------------------
        For i = 0 To UBound(CFGTable.Category)
            ReDim CFGTable.Category(i).condition(init_arrSize)
        Next i
        Dim m_condcnt As Long
        Dim m_pkgheadStr As String
        Dim m_pkgname As String
        Dim m_BinStr As String
        Dim m_bitstrL As String
        Dim m_bitStrM As String
        Dim m_bitStrM2 As String
        Dim m_HexStr As String
        Dim m_hexStr2 As String
        Dim m_bitArr() As Long
        Dim m_dbl As Double

        Dim m_max_msbbit As Long
        Dim m_min_lsbbit As Long
        m_condcnt = 0
        ''''----------------------------------------------------
                
        ''''Then get the following parameter per Header
        M = 0
        Do While (M <= idx_END)
            M = M + 1 ''''Column direction
            n = 0 ''''index and row direction

            Set myCell = myCell_Header.Offset(rowOffset:=0, columnOffset:=(M - 1)) ''''rowOffset MUST be always '0'
            m_cellStr = UCase(Trim(myCell.Value))
            m_lastrow = myCell.End(xlDown).row
            m_lastNCnt = m_lastrow - idx_1stHeader_row
            DebugPrintLog "6-0...input Header =" + m_cellStr + " (" & myCell.Value & "), m=" + CStr(M) + ", LastRow=" + CStr(m_lastrow) + ", LastNCnt=" + CStr(m_lastNCnt)

            Select Case (M)
            Case idx_END
                gB_findCFGCondTable_flag = True
                Exit Do ''''end

            Case idx_Condition
                Do While (n < m_lastNCnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.row & ", Condition Index n=" & n
                    'CFGCond.Category(n).index = n
                    'CFGCond.Category(n).Name = m_cellStr
                    For i = 0 To UBound(CFGTable.Category)
                        CFGTable.Category(i).condition(n).Name = m_cellStr
                    Next i
                    n = n + 1
                Loop
                m_condcnt = n
                For i = 0 To UBound(CFGTable.Category)
                    ReDim Preserve CFGTable.Category(i).condition(m_condcnt - 1)
                Next i
                n = 0

            Case idx_MSB
                ''Do While (n < m_lastNCnt)
                m_max_msbbit = 0
                Do While (n < m_condcnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.row & ", n=" & n
                    ''CFGCond.Category(n).MSBbit = CLng(m_cellStr)
                    m_MSBBit = CLng(m_cellStr)
                    If (m_MSBBit >= m_max_msbbit) Then
                        m_max_msbbit = m_MSBBit
                    End If
                    For i = 0 To UBound(CFGTable.Category)
                        CFGTable.Category(i).condition(n).MSBbit = m_MSBBit
                    Next i
                    n = n + 1
                Loop
                n = 0

            Case idx_LSB
                ''Do While (n < m_lastNCnt)
                m_min_lsbbit = 9999999
                Do While (n < m_condcnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.row & ", n=" & n
                    'CFGCond.Category(n).LSBbit = CLng(m_cellStr)
                    m_LSBbit = CLng(m_cellStr)
                    If (m_LSBbit <= m_min_lsbbit) Then
                        m_min_lsbbit = m_LSBbit
                    End If
                    For i = 0 To UBound(CFGTable.Category)
                        CFGTable.Category(i).condition(n).LSBbit = m_LSBbit
                    Next i
                    n = n + 1
                Loop
                n = 0
                
            Case idx_BitWidth
                ''Do While (n < m_lastNCnt)
                Do While (n < m_condcnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = Trim(myCell.Value)
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.row & ", n=" & n
                    m_bitwidth = CLng(m_cellStr)
                    'CFGCond.Category(n).BitWidth = m_bitwidth
                    For i = 0 To UBound(CFGTable.Category)
                        m_MSBBit = CFGTable.Category(i).condition(n).MSBbit
                        m_LSBbit = CFGTable.Category(i).condition(n).LSBbit
                        CFGTable.Category(i).condition(n).BitWidth = Abs(m_MSBBit - m_LSBbit) + 1
                        ''''do the pre-check here
                        If (m_bitwidth <> Abs(m_MSBBit - m_LSBbit) + 1) Then
                            m_tmpStr = CFGTable.Category(i).condition(n).Name + " BitWidth(" + CStr(m_bitwidth) + " is NOT equal to [MSBBit(" + CStr(m_MSBBit) + ") - LSBBit(" + CStr(m_LSBbit) + ") + 1]"
                            TheExec.AddOutput m_tmpStr
                            TheExec.Datalog.WriteComment m_tmpStr
                            GoTo errHandler
                        End If
                    Next i
                    n = n + 1
                Loop
                n = 0

            Case idx_Stage
                ''Do While (n < m_lastNCnt)
                Do While (n < m_condcnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "6...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.row & ", n=" & n
                    'CFGCond.Category(n).Stage = m_cellStr
                    For i = 0 To UBound(CFGTable.Category)
                        CFGTable.Category(i).condition(n).Stage = m_cellStr
                    Next i
                    n = n + 1
                Loop
                n = 0
            
            ''''-----------------------------------------
            ''''Judge if it's Binary, Hex, or Decimal
            ''''-----------------------------------------
            Case idx_Default, idx_ALL_0, idx_1st_SI, idx_A00 To idx_A12, idx_B00 To idx_B12, idx_C00 To idx_C12, idx_D00 To idx_D12
                ''Do While (n < m_lastNCnt)
                m_pkgheadStr = m_cellStr ''''it presents one of the head "A00"..."A12"
                Do While (n < m_condcnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    DebugPrintLog "7...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.row & ", n=" & n
                    
                    If (InStr(1, m_cellStr, "B") = 1) Then
                        ''''case Binary
                        ''Debug.Print "here is " & InStr(1, m_cellStr, "B")
                        m_BinStr = Replace(m_cellStr, "B", "", 1, 1) ''''remove the first 'B' character
                        m_BinStr = Replace(m_BinStr, "_", "")        ''''remove '_' character, 20171211 update for case "b00_1100"
                        For i = 0 To UBound(CFGTable.Category)
                            m_pkgname = CFGTable.Category(i).pkgName
                            If (m_pkgname = m_pkgheadStr) Then
                                m_bitwidth = CFGTable.Category(i).condition(n).BitWidth
                                m_HexStr = auto_BinStr2HexStr(m_BinStr, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))
                                CFGTable.Category(i).condition(n).HexStr = "0x" + m_HexStr
                                m_bitstrL = StrReverse(m_BinStr) ''''LSB...MSB
                                ReDim m_bitArr(m_bitwidth - 1)
                                For j = 1 To Len(m_bitstrL)
                                    m_bitArr(j - 1) = CLng(Mid(m_bitstrL, j, 1))
                                Next j
                                m_dbl = 0# ''''<MUST>
                                ReDim CFGTable.Category(i).condition(n).BitVal(m_bitwidth - 1)
                                
                                m_bitStrM = "" ''''<MUST> clear
                                For j = 0 To UBound(m_bitArr)
                                    CFGTable.Category(i).condition(n).BitVal(j) = m_bitArr(j)
                                    m_dbl = m_dbl + m_bitArr(j) * CDbl(2 ^ j) ''''use CDBL if bits are over 31bits
                                    m_bitStrM = CStr(m_bitArr(j)) + m_bitStrM
                                Next j
                                CFGTable.Category(i).condition(n).Decimal = m_dbl
                                CFGTable.Category(i).condition(n).BitStrM = m_bitStrM
                                Exit For ''''<MUST> shorten search time
                            End If
                        Next i
                        
                    ElseIf (InStr(1, m_cellStr, "X") = 1 Or InStr(1, m_cellStr, "0X") = 1) Then
                        ''''case Hexadecimal
                        'Debug.Print "here X  is " & InStr(1, m_cellStr, "X")
                        'Debug.Print "here 0X is " & InStr(1, m_cellStr, "0X")
                        If (InStr(1, m_cellStr, "0X") = 1) Then
                            m_HexStr = Replace(m_cellStr, "0X", "", 1, 1) ''''remove '0X' character
                        ElseIf (InStr(1, m_cellStr, "X") = 1) Then
                            m_HexStr = Replace(m_cellStr, "X", "", 1, 1) ''''remove 'X' character
                        End If
                        
                        For i = 0 To UBound(CFGTable.Category)
                            m_pkgname = CFGTable.Category(i).pkgName
                            If (m_pkgname = m_pkgheadStr) Then
                                m_bitwidth = CFGTable.Category(i).condition(n).BitWidth
                                CFGTable.Category(i).condition(n).HexStr = "0x" + m_HexStr
                                m_bitStrM = auto_Hex2BinStr(m_HexStr, m_bitwidth)
                                m_bitstrL = StrReverse(m_bitStrM) ''''LSB...MSB
                                ReDim m_bitArr(m_bitwidth - 1)
                                For j = 1 To Len(m_bitstrL)
                                    m_bitArr(j - 1) = CLng(Mid(m_bitstrL, j, 1))
                                Next j
                                m_dbl = 0# ''''<MUST>
                                ReDim CFGTable.Category(i).condition(n).BitVal(m_bitwidth - 1)
                                For j = 0 To UBound(m_bitArr)
                                    CFGTable.Category(i).condition(n).BitVal(j) = m_bitArr(j)
                                    m_dbl = m_dbl + m_bitArr(j) * CDbl(2 ^ j) ''''use CDBL if bits are over 31bits
                                Next j
                                CFGTable.Category(i).condition(n).Decimal = m_dbl
                                CFGTable.Category(i).condition(n).BitStrM = m_bitStrM
                                Exit For ''''<MUST> shorten search time
                            End If
                        Next i
                    Else
                        ''''check if it's number
                        If (IsNumeric(m_cellStr) = True) Then
                            For i = 0 To UBound(CFGTable.Category)
                                m_pkgname = CFGTable.Category(i).pkgName
                                If (m_pkgname = m_pkgheadStr) Then
                                    m_bitwidth = CFGTable.Category(i).condition(n).BitWidth
                                    CFGTable.Category(i).condition(n).Decimal = CLng(m_cellStr)
                                    m_bitStrM = auto_Dec2Bin_EFuse(CLng(m_cellStr), m_bitwidth, m_bitArr)
                                    m_bitstrL = StrReverse(m_bitStrM) ''''LSB...MSB
                                    m_HexStr = auto_BinStr2HexStr(m_bitStrM, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))
                                    CFGTable.Category(i).condition(n).HexStr = "0x" + m_HexStr
                                    ReDim CFGTable.Category(i).condition(n).BitVal(m_bitwidth - 1)
                                    For j = 0 To UBound(m_bitArr)
                                        CFGTable.Category(i).condition(n).BitVal(j) = m_bitArr(j)
                                    Next j
                                    CFGTable.Category(i).condition(n).BitStrM = m_bitStrM
                                    Exit For ''''<MUST> shorten search time
                                End If
                            Next i
                        Else
                            m_tmpStr = "PKGName: " + m_pkgheadStr + ", its content (" + m_cellStr + ") is NOT a Number."
                            TheExec.AddOutput m_tmpStr
                            TheExec.Datalog.WriteComment m_tmpStr
                            GoTo errHandler
                        End If
                    End If
                    n = n + 1
                Loop
                n = 0

            Case idx_CommentA
                ''Do While (n < m_lastNCnt)
                Do While (n < m_condcnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    If (m_cellStr = "") Then m_cellStr = "NA" ''''<MUST>
                    DebugPrintLog "8...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.row & ", n=" & n
                    For i = 0 To UBound(CFGTable.Category)
                        m_pkgname = CFGTable.Category(i).pkgName
                        If (m_pkgname Like "A##") Then
                            CFGTable.Category(i).condition(n).comment = m_cellStr
                        End If
                    Next i
                    n = n + 1
                Loop
                n = 0

            Case idx_CommentB
                ''Do While (n < m_lastNCnt)
                Do While (n < m_condcnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    If (m_cellStr = "") Then m_cellStr = "NA" ''''<MUST>
                    DebugPrintLog "9...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.row & ", n=" & n
                    For i = 0 To UBound(CFGTable.Category)
                        m_pkgname = CFGTable.Category(i).pkgName
                        If (m_pkgname Like "B##") Then
                            CFGTable.Category(i).condition(n).comment = m_cellStr
                        End If
                    Next i
                    n = n + 1
                Loop
                n = 0

            Case idx_CommentC, idx_CommentD
                ''Do While (n < m_lastNCnt)
                Do While (n < m_condcnt)
                    Set myCell = myCell.Offset(rowOffset:=1, columnOffset:=0)
                    m_cellStr = UCase(Trim(myCell.Value))
                    If (m_cellStr = "") Then m_cellStr = "NA" ''''<MUST>
                    DebugPrintLog "9...input m_cellStr=" + m_cellStr + " (" & myCell.Value & "), Row=" & myCell.row & ", n=" & n
                    For i = 0 To UBound(CFGTable.Category)
                        m_pkgname = CFGTable.Category(i).pkgName
                        If (m_pkgname Like "C##" Or m_pkgname Like "D##") Then
                            CFGTable.Category(i).condition(n).comment = m_cellStr
                        End If
                    Next i
                    n = n + 1
                Loop
                n = 0
            Case Else
                DebugPrintLog "xxx...No Assigned Column(" & M & ") !!!"
            End Select

        Loop ''''end of Do While (m <= idx_END)
        
    End If ''''end of If (find_AllHeader) Then

    ''''-----------------------------------------------------------------------------
    ''''PostProcess to get the following contents
    ''''-----------------------------------------------------------------------------
    ''''Here Category(i) is PKG which means A00...A12, B00...B12, C00...C12, D00...D12
    ''''CFGTable.Category(i).BitStrM
    ''''                    .BitStr() ...BitStr String type
    ''''                    .BitVal() ...BitVal Long type
    ''''                    .BitStrM_byStage  ...BitStr String type by Stage
    ''''                    .BitVal_byStage() ...BitVal Long type by Stage
    ''''                    .BitStr_byStage() ...BitStr String type by Stage
    ''''-----------------------------------------------------------------------------
    '''' Condition Data Structure
    ''''-----------------------------------------------------------------------------
    '''' Here Category(i).Condition(j)
    ''''      is Condition Content of the specific PKG A00...A12, B00...B12, C00...C12, D00...D12
    ''''CFGTable.Category(i).Condition(j).Name
    ''''                                 .Stage
    ''''                                 .LSBbit
    ''''                                 .MSBbit
    ''''                                 .BitWidth
    ''''                                 .HexStr     (with the prefix 0x)
    ''''                                 .Decimal
    ''''                                 .BitStrM
    ''''                                 .BitVal()
    ''''                                 .Comment
    ''''-----------------------------------------------------------------------------
    '''' Here Category(i).Cate32bit(k)
    ''''      is per 32bits Content of the specific PKG A00...A12, B00...B12, C00...C12, D00...D12
    ''''CFGTable.Category(i).Cate32bit(k).Name
    ''''                                 .Stage
    ''''                                 .LSBbit
    ''''                                 .MSBbit
    ''''                                 .BitWidth
    ''''                                 .HexStr     (with the prefix 0x)
    ''''                                 .Decimal
    ''''                                 .BitStrM
    ''''                                 .BitVal()
    ''''                                 ---------------
    ''''                                 .HexStr_byStage   (with the prefix 0x)
    ''''                                 .Decimal_byStage
    ''''                                 .BitStrM_byStage  ...BitStr String type by Stage
    ''''                                 .BitVal_byStage() ...BitVal Long type by Stage
    ''''                                 .Comment
    ''''-----------------------------------------------------------------------------
    DebugPrintLog "99...Max MSBbit=" & m_max_msbbit & ", min LSBbit=" & m_min_lsbbit


'''' if (False) Then gS_JobName = "ft3" ''''debug and simulation

    Dim showPrint As Boolean
    Dim m_stage As String
    Dim m_totalbits As Long
    Dim m_bitval As Long
    Dim m_tmpStr1 As String
    Dim m_tmpStr2 As String

    showPrint = False ''''True, debug purpose
    m_totalbits = Abs(m_max_msbbit - m_min_lsbbit) + 1
    
    ''''20180917 update to get the total bits for the non-continuous bits sequence
    Dim m_bitcnt As Long
    ReDim gL_CFG_Cond_BitIndex(m_totalbits - 1)
    m_bitcnt = 0
    For i = m_min_lsbbit To m_max_msbbit
        For j = 0 To UBound(CFGTable.Category(0).condition)
            m_LSBbit = CFGTable.Category(0).condition(j).LSBbit
            m_MSBBit = CFGTable.Category(0).condition(j).MSBbit
            If (i >= m_LSBbit And i <= m_MSBBit) Then
                gL_CFG_Cond_BitIndex(m_bitcnt) = i
                m_bitcnt = m_bitcnt + 1
                Exit For ''''escape j-loop
            End If
        Next j
    Next i
    
    ''''update m_totalbits to the real effective bits
    m_totalbits = m_bitcnt
    ReDim Preserve gL_CFG_Cond_BitIndex(m_totalbits - 1)

    ''''20170911 update
    gL_CFG_Cond_allBitWidth = m_totalbits
    gL_CFG_Cond_min_lsbbit = m_min_lsbbit
    gL_CFG_Cond_max_msbbit = m_max_msbbit

    ''''20180917 update
    Dim kcnt As Long
    Dim kcnt_s As Long

    For i = 0 To UBound(CFGTable.Category)
        ''''initialize ---------------------------------------------
        ReDim CFGTable.Category(i).Cate32bit(m_totalbits - 1)
        ''''--------------------------------------------------------
        ReDim CFGTable.Category(i).BitVal(m_totalbits - 1)
        ReDim CFGTable.Category(i).BitStr(m_totalbits - 1)
        ReDim CFGTable.Category(i).BitVal_byStage(m_totalbits - 1)
        ReDim CFGTable.Category(i).BitStr_byStage(m_totalbits - 1)

        ''''--------------------------------------------------------
        m_pkgname = CFGTable.Category(i).pkgName
        kcnt = 0
        kcnt_s = 0
        For j = 0 To UBound(CFGTable.Category(i).condition)
            m_stage = CFGTable.Category(i).condition(j).Stage
            m_LSBbit = CFGTable.Category(i).condition(j).LSBbit
            m_MSBBit = CFGTable.Category(i).condition(j).MSBbit
            
            If (gB_eFuse_CFG_Cond_FTF_done_Flag = True) Then ''''20170923 add
                For n = m_LSBbit To m_MSBBit
                    M = n - m_LSBbit
                    m_bitval = CFGTable.Category(i).condition(j).BitVal(M)
                    CFGTable.Category(i).BitVal(kcnt) = m_bitval
                    CFGTable.Category(i).BitStr(kcnt) = CStr(m_bitval)
                    kcnt = kcnt + 1
                Next n
            ElseIf (checkJob_less_Stage_Sequence(m_stage) = False) Then ''''Job >= m_Stage
                For n = m_LSBbit To m_MSBBit
                    M = n - m_LSBbit
                    m_bitval = CFGTable.Category(i).condition(j).BitVal(M)
                    CFGTable.Category(i).BitVal(kcnt) = m_bitval
                    CFGTable.Category(i).BitStr(kcnt) = CStr(m_bitval)
                    kcnt = kcnt + 1
                Next n
            Else
                ''''others keep zero
                For n = m_LSBbit To m_MSBBit
                    m_bitval = 0
                    CFGTable.Category(i).BitVal(kcnt) = m_bitval
                    CFGTable.Category(i).BitStr(kcnt) = CStr(m_bitval)
                    kcnt = kcnt + 1
                Next n
            End If
            
            ''''<Important>
            ''''Here is used for the programming by Stage
            ''''set condition with stage=Job only, others keep zero
            If (LCase(m_stage) = gS_JobName) Then ''''case CP1 or FT3 (FTF)
                For n = m_LSBbit To m_MSBBit
                    M = n - m_LSBbit
                    m_bitval = CFGTable.Category(i).condition(j).BitVal(M)
                    CFGTable.Category(i).BitVal_byStage(kcnt_s) = m_bitval
                    CFGTable.Category(i).BitStr_byStage(kcnt_s) = CStr(m_bitval)
                    kcnt_s = kcnt_s + 1
                Next n
            Else
                ''''others keep zero
                For n = m_LSBbit To m_MSBBit
                    m_bitval = 0
                    CFGTable.Category(i).BitVal_byStage(kcnt_s) = m_bitval
                    CFGTable.Category(i).BitStr_byStage(kcnt_s) = CStr(m_bitval)
                    kcnt_s = kcnt_s + 1
                Next n
            End If
        Next j
        
        ''''Get BitStrM per PKG Name
        m_bitStrM = ""
        m_bitStrM2 = ""
        m_bitwidth = m_totalbits
        For n = 0 To m_totalbits - 1
            m_bitStrM = CFGTable.Category(i).BitStr(n) + m_bitStrM
            m_bitStrM2 = CFGTable.Category(i).BitStr_byStage(n) + m_bitStrM2
        Next n
        CFGTable.Category(i).BitStrM = m_bitStrM ''''<MUST>
        m_HexStr = "0x" + auto_BinStr2HexStr(m_bitStrM, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))

        CFGTable.Category(i).BitStrM_byStage = m_bitStrM2 ''''<MUST>
        m_hexStr2 = "0x" + auto_BinStr2HexStr(m_bitStrM2, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))

''        If (showPrint) Then
''            TheExec.Datalog.WriteComment UCase(gS_JobName) + " PKG=" + m_pkgname + " Programming "
''            m_tmpStr = "HexStr[" & Format(m_max_msbbit, "000") & ":" & Format(m_min_lsbbit, "000") & "]=" + m_hexStr2
''            TheExec.Datalog.WriteComment Space(5) + m_tmpStr
''            m_tmpStr = "BitStr[" & Format(m_max_msbbit, "000") & ":" & Format(m_min_lsbbit, "000") & "]=" + m_bitStrM2
''            TheExec.Datalog.WriteComment Space(5) + m_tmpStr ''''& vbCrLf
''            TheExec.Datalog.WriteComment Space(5) + "----------------------------------------------------------------------------"
''        End If

        ''''Get BitStrM per 32bits per PKG Name
        m_bitStrM = ""
        m_bitwidth = 32
        Dim mm_msbbit As Long
        Dim mm_lsbbit As Long
        If ((m_totalbits Mod m_bitwidth) <> 0) Then TheExec.Datalog.WriteComment "<WARNING> Total Bits are NOT multiple of " & m_bitwidth

        ''''=======================================================
        ''''display from [minbit......maxbit]
        Dim m_exist_flag As Boolean
        m_exist_flag = False
        
        k = 0 ''''<MUST>
        m_bitcnt = 0
        For M = m_min_lsbbit To m_max_msbbit
            m_bitStrM = "" ''''<MUST>
            m_HexStr = ""  ''''<MUST>
            m_tmpStr = ""  ''''<MUST>

            mm_lsbbit = M
            mm_msbbit = (M + m_bitwidth) - 1
            
            m_exist_flag = False ''''<MUST>
            ''''here check if M exist in the condition bits
            For j = 0 To UBound(CFGTable.Category(0).condition)
                m_LSBbit = CFGTable.Category(0).condition(j).LSBbit
                m_MSBBit = CFGTable.Category(0).condition(j).MSBbit
                If (M >= m_LSBbit And M <= m_MSBBit) Then
                    m_exist_flag = True
                    Exit For ''''escape j-loop
                End If
            Next j

            If (m_exist_flag = True) Then
                CFGTable.Category(i).Cate32bit(k).MSBbit = mm_msbbit
                CFGTable.Category(i).Cate32bit(k).LSBbit = mm_lsbbit
                CFGTable.Category(i).Cate32bit(k).BitWidth = m_bitwidth
                m_tmpStr = "CFG_Condition_" + m_pkgname + "[" + Format(mm_msbbit, "000") + ":" + Format(mm_lsbbit, "000") + "]"
                CFGTable.Category(i).Cate32bit(k).comment = m_tmpStr
                
                m_tmpStr = ""  ''''<MUST>
                For j = 0 To UBound(CFGTable.Category(0).condition)
                    m_stage = CFGTable.Category(0).condition(j).Stage
                    m_LSBbit = CFGTable.Category(0).condition(j).LSBbit '- m_offset
                    m_MSBBit = CFGTable.Category(0).condition(j).MSBbit '- m_offset
    
                    If (m_MSBBit = mm_msbbit) Then m_tmpStr1 = CFGTable.Category(0).condition(j).Name
                    If (m_LSBbit = mm_lsbbit) Then m_tmpStr2 = CFGTable.Category(0).condition(j).Name
                    If ((m_MSBBit <= mm_msbbit) And (m_LSBbit >= mm_lsbbit)) Then m_tmpStr = m_tmpStr + "," + m_stage
                Next j
                CFGTable.Category(i).Cate32bit(k).Stage = Replace(m_tmpStr, ",", "", 1, 1)
                m_tmpStr = "" ''''reuse variable
                m_tmpStr = m_tmpStr1 + Replace(UCase(m_tmpStr2), UCase("CFG_Condition"), "", 1, 1)
                If (InStr(1, UCase(m_tmpStr), UCase("CFG_Condition"), vbTextCompare) = 0) Then
                    m_tmpStr = "CFG_Condition_" + CStr(k)
                End If
                CFGTable.Category(i).Cate32bit(k).Name = m_tmpStr
    
                m_bitStrM = "" ''''<MUST>
                m_bitStrM2 = "" ''''<MUST>
                ReDim CFGTable.Category(i).Cate32bit(k).BitVal(m_bitwidth - 1) ''''<MUST>
                ReDim CFGTable.Category(i).Cate32bit(k).BitVal_byStage(m_bitwidth - 1) ''''<MUST>

                ''''was mm_lsbbit <= m_bitcnt
                For n = 0 To m_bitwidth - 1
                    m_bitStrM = CFGTable.Category(i).BitStr(m_bitcnt + n) + m_bitStrM
                    m_bitStrM2 = CFGTable.Category(i).BitStr_byStage(m_bitcnt + n) + m_bitStrM2
                    ''''CFGTable.Category(i).Cate32bit(k).BitVal(0) is the LSB bit
                    CFGTable.Category(i).Cate32bit(k).BitVal(n) = CFGTable.Category(i).BitVal(m_bitcnt + n)
                    CFGTable.Category(i).Cate32bit(k).BitVal_byStage(n) = CFGTable.Category(i).BitVal_byStage(m_bitcnt + n)
                Next n
                
                m_HexStr = "0x" + auto_BinStr2HexStr(m_bitStrM, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))
                CFGTable.Category(i).Cate32bit(k).BitStrM = m_bitStrM
                CFGTable.Category(i).Cate32bit(k).HexStr = m_HexStr
                CFGTable.Category(i).Cate32bit(k).Decimal = auto_HexStr2Value(m_HexStr)
              
                ''''by Stage case
                m_hexStr2 = "0x" + auto_BinStr2HexStr(m_bitStrM2, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))
                CFGTable.Category(i).Cate32bit(k).BitStrM_byStage = m_bitStrM2
                CFGTable.Category(i).Cate32bit(k).HexStr_byStage = m_hexStr2
                CFGTable.Category(i).Cate32bit(k).Decimal_byStage = auto_HexStr2Value(m_hexStr2)
''                If (showPrint) Then
''                    m_tmpStr = FormatNumeric(CFGTable.Category(i).Cate32bit(k).Name, -20) + _
''                               "[" + Format(mm_msbbit, "000") + ":" + Format(mm_lsbbit, "000") & "]=" + _
''                               m_bitStrM2 + "=" + m_hexStr2 + "=" + CStr(CFGTable.Category(i).Cate32bit(k).Decimal_byStage)
''                    TheExec.Datalog.WriteComment Space(5) + m_tmpStr
''                End If
                M = mm_msbbit ''''<MUST>
                m_bitcnt = m_bitcnt + m_bitwidth
                k = k + 1 ''''<MUST>
            End If
        Next M
        ReDim Preserve CFGTable.Category(i).Cate32bit(k - 1)
        ''If (showPrint) Then TheExec.Datalog.WriteComment ""
    Next i

    ''TheExec.Datalog.WriteComment "-----------------------------------------------------------------------------------------------------------------"
    ''TheExec.Datalog.WriteComment "-----------------------------------------------------------------------------------------------------------------"
    ''''20170719 update
    If (showPrint) Then
        TheExec.Datalog.WriteComment ""
        For i = 0 To UBound(CFGTable.Category)
            m_pkgname = CFGTable.Category(i).pkgName
            
            ''''Here it means that the CFG_Condition result (shoule be) programming in the specific Category (Axx,Bxx)
            m_bitwidth = m_totalbits
            m_bitStrM = CFGTable.Category(i).BitStrM_byStage
            m_HexStr = "0x" + auto_BinStr2HexStr(m_bitStrM, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))

            If (gB_eFuse_CFG_Cond_FTF_done_Flag = False) Then ''''20170923 add
                TheExec.Datalog.WriteComment UCase(gS_JobName) + " PKG=" + m_pkgname + " Programming "
''                m_tmpStr = "HexStr[" & Format(m_max_msbbit, "0000") & ":" & Format(m_min_lsbbit, "000") & "]=" + m_hexStr
''                TheExec.DataLog.WriteComment Space(5) + m_tmpStr
''                m_tmpStr = "BitStr[" & Format(m_max_msbbit, "0000") & ":" & Format(m_min_lsbbit, "000") & "]=" + m_bitStrM
''                TheExec.DataLog.WriteComment Space(5) + m_tmpStr ''''& vbCrLf
                TheExec.Datalog.WriteComment Space(5) + "----------------------------------------------------------------------------"

                For k = 0 To UBound(CFGTable.Category(i).Cate32bit)
                    mm_lsbbit = CFGTable.Category(i).Cate32bit(k).LSBbit
                    mm_msbbit = CFGTable.Category(i).Cate32bit(k).MSBbit
                    m_bitStrM = CFGTable.Category(i).Cate32bit(k).BitStrM_byStage
                    m_HexStr = CFGTable.Category(i).Cate32bit(k).HexStr_byStage
                    m_tmpStr = FormatNumeric(CFGTable.Category(i).Cate32bit(k).Name, -20) + _
                               "[" + Format(mm_msbbit, "000") + ":" + Format(mm_lsbbit, "000") & "]=" + _
                               m_bitStrM + "=" + m_HexStr + "=" + CStr(CFGTable.Category(i).Cate32bit(k).Decimal_byStage)
                    TheExec.Datalog.WriteComment Space(5) + m_tmpStr
                Next k
                TheExec.Datalog.WriteComment ""
            End If
            
            ''''Here it means that the CFG_Condition result (shoule be) after fused/blown
            m_bitwidth = m_totalbits
            m_bitStrM = CFGTable.Category(i).BitStrM
            m_HexStr = "0x" + auto_BinStr2HexStr(m_bitStrM, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))
            TheExec.Datalog.WriteComment Space(4) + "PKG=" + m_pkgname + " After Fused (should be like as)"
''            m_tmpStr = "HexStr[" & Format(m_max_msbbit, "0000") & ":" & Format(m_min_lsbbit, "000") & "]=" + m_hexStr
''            TheExec.DataLog.WriteComment Space(5) + m_tmpStr
''            m_tmpStr = "BitStr[" & Format(m_max_msbbit, "0000") & ":" & Format(m_min_lsbbit, "000") & "]=" + m_bitStrM
''            TheExec.DataLog.WriteComment Space(5) + m_tmpStr ''''& vbCrLf
            TheExec.Datalog.WriteComment Space(5) + "----------------------------------------------------------------------------"
    
            For k = 0 To UBound(CFGTable.Category(i).Cate32bit)
                mm_lsbbit = CFGTable.Category(i).Cate32bit(k).LSBbit
                mm_msbbit = CFGTable.Category(i).Cate32bit(k).MSBbit
                m_bitStrM = CFGTable.Category(i).Cate32bit(k).BitStrM
                m_HexStr = CFGTable.Category(i).Cate32bit(k).HexStr
                m_tmpStr = FormatNumeric(CFGTable.Category(i).Cate32bit(k).Name, -20) + _
                           "[" + Format(mm_msbbit, "000") + ":" + Format(mm_lsbbit, "000") & "]=" + _
                           m_bitStrM + "=" + m_HexStr + "=" + CStr(CFGTable.Category(i).Cate32bit(k).Decimal)
                TheExec.Datalog.WriteComment Space(5) + m_tmpStr
            Next k
            TheExec.Datalog.WriteComment ""
        Next i
    End If

    ''''201811XX
    ''''build up the global Dictionary to speed up the Index search
    For j = 0 To UBound(CFGTable.Category)
        m_tmpStr1 = CFGTable.Category(j).pkgName
        Call eFuse_AddStoredIndex(eFuse_CFGTab, m_tmpStr1, j)
    Next j
    
    For j = 0 To UBound(CFGTable.Category(0).condition)
        m_tmpStr2 = CFGTable.Category(0).condition(j).Name
        Call eFuse_AddStoredIndex(eFuse_CFGCond, m_tmpStr2, j)
    Next j
    
    If (showPrint) Then ''''debug purpose
        For j = 0 To UBound(CFGTable.Category)
            m_tmpStr1 = CFGTable.Category(j).pkgName
            k = eFuse_GetStoredIndex(eFuse_CFGTab, m_tmpStr1)
            TheExec.Datalog.WriteComment m_tmpStr1 & " Index = " & k
        Next j
        
        For j = 0 To UBound(CFGTable.Category(0).condition)
            m_tmpStr2 = CFGTable.Category(0).condition(j).Name
            k = eFuse_GetStoredIndex(eFuse_CFGCond, m_tmpStr2)
            TheExec.Datalog.WriteComment m_tmpStr2 & " Index = " & k
        Next j
    End If
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function auto_Hex2BinStr(ByVal HexStr As String, Optional BitWidth As Long = 0) As String
                                                                                                                         
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Hex2BinStr"
    
    Dim i As Long
    Dim PerChar As String
    Dim BinStr As String
    Dim j As Long
    Dim DecodeBin As String
    Dim MyArray() As Variant
    Dim myArrayBin() As Variant
    Dim m_len As Long

    HexStr = UCase(HexStr)
    If (InStr(1, HexStr, "X") = 1) Then
        HexStr = Replace(HexStr, "X", "", 1, 1)
    ElseIf (InStr(1, HexStr, "0X") = 1) Then
        HexStr = Replace(HexStr, "0X", "", 1, 1)
    End If

    MyArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                    "A", "B", "C", "D", "E", "F")

    myArrayBin = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", "1000", "1001", _
                       "1010", "1011", "1100", "1101", "1110", "1111")

    BinStr = ""
    For i = 1 To Len(HexStr)
        PerChar = Mid(HexStr, i, 1)
        'One-to-One mapping, myarray() mappping to myarraybin()
        For j = 0 To UBound(MyArray)
            If (PerChar = MyArray(j)) Then
               DecodeBin = myArrayBin(j)
               Exit For
            End If
        Next j
        BinStr = BinStr + DecodeBin
    Next i

    m_len = Len(BinStr)
    If (BitWidth <> 0) Then
        If (m_len > BitWidth) Then
            For i = 1 To (m_len - BitWidth)
                PerChar = Mid(BinStr, 1, 1)
                If (PerChar = "0") Then
                    BinStr = Replace(BinStr, "0", "", 1, 1)
                Else
                    TheExec.AddOutput "<WARNING> " + funcName + ":: Effect BinStr(" + BinStr + ") Length(" & Len(BinStr) & ") > BitWidth(" & BitWidth & ")"
                    TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: Effect BinStr(" + BinStr + ") Length(" & Len(BinStr) & ") > BitWidth(" & BitWidth & ")"
                End If
            Next i
        ElseIf (m_len < BitWidth) Then
            For i = 1 To (BitWidth - m_len)
                BinStr = "0" + BinStr
            Next i
        End If
    End If
    auto_Hex2BinStr = BinStr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function

''''20170630 update
Public Function CFGCondTabIndex(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CFGCondTabIndex"

    Dim i As Long
    Dim match_Flag As Boolean
    match_Flag = False
    
    If (True) Then
        Dim m_keyname As String
        m_keyname = myStr
        CFGCondTabIndex = eFuse_GetStoredIndex(eFuse_CFGCond, m_keyname)
        If (CFGCondTabIndex >= 0) Then match_Flag = True
    Else
        For i = 0 To UBound(CFGTable.Category(0).condition)
            If (UCase(myStr) = UCase(CFGTable.Category(0).condition(i).Name)) Then
                CFGCondTabIndex = i
                match_Flag = True
                Exit For
            End If
        Next i
    End If

    If (match_Flag = False) Then
        CFGCondTabIndex = -1
        PrintDataLog funcName + ":: <" + myStr + ">, it's NOT existed in the Category.Condition."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20170813, New
''''process Value to Hex String.
''''20171211, support Value=binary to HexStr
''''201811XX update
Public Function auto_Value2HexStr(ByVal Value As Variant, Optional BitWidth As Long = 0, Optional showPrint As Boolean = False) As String

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Value2HexStr"

    ''''-------------------------------------------------------------------------
    ''''Example, it could support up to (2^1024 -1)
    ''''-------------------------------------------------------------------------
    ''''Call auto_Value2HexStr("0x7FFFFFFF") = 0x7FFFFFFF
    ''''Call auto_Value2HexStr(2147483647)   = 0x7FFFFFFF
    ''''Call auto_Value2HexStr(9876543210)   = 0x24CB016EA
    ''''Call auto_Value2HexStr("9876543210", 50) = 0x000024CB016EA
    ''''Call auto_Value2HexStr("b00111110110")   = 0x1F6
    ''''Call auto_Value2HexStr("b001_1111_0110") = 0x1F6
    ''''-------------------------------------------------------------------------
    Dim i As Long
    Dim m_HexStr As String
    Dim m_tmpStr As String
    Dim m_dbl As Double
    Dim m_dbl_f As Double
    Dim m_dbl_g As Long
    
    Dim m_hexStr_raw As String
    Dim m_displayHexLen  As Long
    Dim m_hexlen As Long
    Dim m_bw_q4 As Long
    ''''debug try value = CDbl(2 ^ 1023 + 2 ^ 1022)

    If (IsNumeric(CStr(Value)) = True) Then
        m_dbl = CDbl(Value)
        If (m_dbl <= CLng("&H7FFFFFFF")) Then ''''CLng("&H7FFFFFFF")=2147483647
            m_HexStr = "0x" + Hex(Value)
        ElseIf (m_dbl <= 9.00719925474099E+15) Then ''''0x1FFFFFFFFFFFFF=9007199254740991#
            ''''<NOTICE> Only support up to [0x1FFFFFFFFFFFFF=9007199254740991] then can NOT get the correct quotient/remainder
            ''''It's only 53bits
            ''m_tmpStr = funcName + ":: Input Value (" + CStr(value) + ") is Over Long Value limit (0x7FFFFFFF=2147483647)."
            ''TheExec.Datalog.WriteComment m_tmpStr
            m_HexStr = ""
            Do
                m_dbl_f = Floor(m_dbl / 16)
                m_dbl_g = m_dbl - (m_dbl_f * 16#)
                m_HexStr = Hex(CStr(m_dbl_g)) + m_HexStr

                m_tmpStr = "m_dbl=" + FormatNumeric(m_dbl, -15) + ", m_dbl_f=" + FormatNumeric(m_dbl_f, -15) + ", m_dbl_g=" + FormatNumeric(m_dbl_g, -15) + ", m_hexStr=" + m_HexStr
                If (showPrint) Then TheExec.Datalog.WriteComment m_tmpStr
                
                ''''for next process
                m_dbl = m_dbl_f
            Loop While (m_dbl > 16#)

            m_HexStr = "0x" + Hex(CStr(m_dbl_f)) + m_HexStr
        Else
            TheExec.AddOutput "<Error> " + funcName + ":: Can NOT do the conversion of m_dbl=" & m_dbl
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: Can NOT do the conversion of m_dbl=" & m_dbl
            GoTo errHandler
        End If

    ElseIf (auto_isHexString(CStr(Value)) = True) Then
        m_HexStr = CStr(Value)
        m_HexStr = Replace(m_HexStr, "_", "") ''''20171103 update, for case like as 0x1234_ABCD_EFFE
        ''''case: "x012ABC" to "0x012ABC"
        If (InStr(1, m_HexStr, "x", vbTextCompare) = 1) Then m_HexStr = Replace(m_HexStr, "x", "0x", 1, 1, vbTextCompare)
    
    ElseIf (auto_isBinaryString(CStr(Value))) Then ''''20171211 add
        Dim m_BinStr As String
        m_BinStr = Replace(UCase(CStr(Value)), "B", "", 1, 1) ''''remove the first "B" character
        m_BinStr = Replace(m_BinStr, "_", "")                 ''''for case like as b0001_0101
        
        ''''convert to Hex with the prefix "0x"
        m_HexStr = "0x" + auto_BinStr2HexStr(m_BinStr, 1)

    Else
        m_tmpStr = LCase(CStr(Value))
        If (m_tmpStr = "bincut" Or m_tmpStr = "na" Or m_tmpStr = "n/a") Then
            m_HexStr = "0x0"
        Else
            m_tmpStr = "<Error> " + funcName + ":: Input Value (" + CStr(Value) + ") is NOT a Numeric, Binary or Hex number."
            TheExec.AddOutput m_tmpStr
            TheExec.Datalog.WriteComment m_tmpStr
            GoTo errHandler
        End If
    End If

    ''''process the HexStr with the related bitwidth
    If (BitWidth > 0) Then
        m_hexStr_raw = Replace(m_HexStr, "0x", "", 1, 1, vbTextCompare)
        m_hexlen = Len(m_hexStr_raw) ''''ignore first 2 chars '0x'
        m_bw_q4 = BitWidth \ 4 '''get quotient (devide by 4)
        m_displayHexLen = IIf(BitWidth - m_bw_q4 * 4 = 0, m_bw_q4, m_bw_q4 + 1)
        
        If (m_hexlen < m_displayHexLen) Then
            For i = 1 To (m_displayHexLen - m_hexlen)
                m_hexStr_raw = "0" + m_hexStr_raw
            Next i
            m_HexStr = "0x" + m_hexStr_raw
        Else
            ''''do Nothing here
        End If
    End If

    If (showPrint) Then
        m_tmpStr = funcName + ":: m_hexStr = " + m_HexStr + " (" + CStr(Value) + ")"
        TheExec.Datalog.WriteComment m_tmpStr
    End If
    
    auto_Value2HexStr = m_HexStr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20170813 New
''''It's used to do the String compare for the limit.
''''<MUST> the compare string must be the same type, all Hex String or all Binary String.
Public Function auto_TestStringLimit(testStr As String, lolmtStr As String, hilmtStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_TestStringLimit"

    Dim m_result As Long
    Dim m_cmphiVal As Long
    Dim m_cmploVal As Long
    
    ''''<MUST>
    testStr = UCase(testStr)
    lolmtStr = UCase(lolmtStr)
    hilmtStr = UCase(hilmtStr)
    ''''------------------------------------------
    ''''compare with lolmt, hilmt
    ''''return -1 means less
    ''''return  0 means equal
    ''''return  1 means large
    ''''------------------------------------------
    m_cmploVal = StrComp(testStr, lolmtStr, vbBinaryCompare)
    m_cmphiVal = StrComp(testStr, hilmtStr, vbBinaryCompare)
    
    If (m_cmploVal = 0 And m_cmphiVal = 0) Then
        '''' =lolimt and =hilimt, pass
        m_result = 1
    ElseIf (m_cmploVal = 1 And m_cmphiVal <> 1) Then
        '''' >lolimt and <=hilimt, pass
        m_result = 1
    ElseIf (m_cmploVal = -1) Then
        '''' < lolimt, fail
        m_result = 0
    ElseIf (m_cmphiVal = 1) Then
        '''' > hilimt, fail
        m_result = 0
    Else
        m_result = 0
    End If
    
    ''''Pass: m_Result=1
    ''''Fail: m_Result=0
    auto_TestStringLimit = m_result

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20170911 Add
Public Function auto_display_CFG_Cond_Table_by_PKGName(ByVal pkgName As String, Optional showPrint As Boolean = True) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_display_CFG_Cond_Table_by_PKGName"

    auto_display_CFG_Cond_Table_by_PKGName = 0
    
    Dim i As Long, k As Long
    Dim m_totalbits As Long
    Dim m_max_msbbit As Long
    Dim m_min_lsbbit As Long
    Dim m_bitwidth As Long
    Dim mm_lsbbit As Long
    Dim mm_msbbit As Long
    Dim m_pkgname As String
    Dim m_bitStrM As String
    Dim m_HexStr As String
    Dim m_tmpStr As String
    
    pkgName = UCase(Trim(pkgName))
    
    If (gB_findCFGCondTable_flag = False) Then Exit Function
    
    If (showPrint) Then
        m_min_lsbbit = gL_CFG_Cond_min_lsbbit
        m_max_msbbit = gL_CFG_Cond_max_msbbit

        TheExec.Datalog.WriteComment ""
        For i = 0 To UBound(CFGTable.Category)
            m_pkgname = UCase(Trim(CFGTable.Category(i).pkgName))
            If (pkgName = m_pkgname) Then
                ''''Here it means that the CFG_Condition result (shoule be) programming in the specific Category (Axx,Bxx)
                m_bitwidth = gL_CFG_Cond_allBitWidth
                m_bitStrM = CFGTable.Category(i).BitStrM_byStage
                m_HexStr = "0x" + auto_BinStr2HexStr(m_bitStrM, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))
    
                If (gB_eFuse_CFG_Cond_FTF_done_Flag = False) Then ''''20170923 add
                    TheExec.Datalog.WriteComment UCase(gS_JobName) + " PKG=" + m_pkgname + " Programming "
                    m_tmpStr = "HexStr[" & Format(m_max_msbbit, "000") & ":" & Format(m_min_lsbbit, "000") & "]=" + m_HexStr
                    TheExec.Datalog.WriteComment Space(5) + m_tmpStr
                    m_tmpStr = "BitStr[" & Format(m_max_msbbit, "000") & ":" & Format(m_min_lsbbit, "000") & "]=" + m_bitStrM
                    TheExec.Datalog.WriteComment Space(5) + m_tmpStr ''''& vbCrLf
                    TheExec.Datalog.WriteComment Space(5) + "----------------------------------------------------------------------------"

                    For k = 0 To UBound(CFGTable.Category(i).Cate32bit)
                        mm_lsbbit = CFGTable.Category(i).Cate32bit(k).LSBbit
                        mm_msbbit = CFGTable.Category(i).Cate32bit(k).MSBbit
                        m_bitStrM = CFGTable.Category(i).Cate32bit(k).BitStrM_byStage
                        m_HexStr = CFGTable.Category(i).Cate32bit(k).HexStr_byStage
                        m_tmpStr = FormatNumeric(CFGTable.Category(i).Cate32bit(k).Name, -20) + _
                                   "[" + Format(mm_msbbit, "000") + ":" + Format(mm_lsbbit, "000") & "]=" + _
                                   m_bitStrM + "=" + m_HexStr + "=" + CStr(CFGTable.Category(i).Cate32bit(k).Decimal_byStage)
                        TheExec.Datalog.WriteComment Space(5) + m_tmpStr
                    Next k
                    TheExec.Datalog.WriteComment ""
                End If

                ''''Here it means that the CFG_Condition result (shoule be) after fused/blown
                m_bitwidth = gL_CFG_Cond_allBitWidth
                m_bitStrM = CFGTable.Category(i).BitStrM
                m_HexStr = "0x" + auto_BinStr2HexStr(m_bitStrM, IIf((m_bitwidth Mod 4) = 0, m_bitwidth \ 4, 1 + (m_bitwidth \ 4)))
                TheExec.Datalog.WriteComment Space(4) + "PKG=" + m_pkgname + " After Fused (should be like as)"
                m_tmpStr = "HexStr[" & Format(m_max_msbbit, "000") & ":" & Format(m_min_lsbbit, "000") & "]=" + m_HexStr
                TheExec.Datalog.WriteComment Space(5) + m_tmpStr
                m_tmpStr = "BitStr[" & Format(m_max_msbbit, "000") & ":" & Format(m_min_lsbbit, "000") & "]=" + m_bitStrM
                TheExec.Datalog.WriteComment Space(5) + m_tmpStr ''''& vbCrLf
                TheExec.Datalog.WriteComment Space(5) + "----------------------------------------------------------------------------"
        
                For k = 0 To UBound(CFGTable.Category(i).Cate32bit)
                    mm_lsbbit = CFGTable.Category(i).Cate32bit(k).LSBbit
                    mm_msbbit = CFGTable.Category(i).Cate32bit(k).MSBbit
                    m_bitStrM = CFGTable.Category(i).Cate32bit(k).BitStrM
                    m_HexStr = CFGTable.Category(i).Cate32bit(k).HexStr
                    m_tmpStr = FormatNumeric(CFGTable.Category(i).Cate32bit(k).Name, -20) + _
                               "[" + Format(mm_msbbit, "000") + ":" + Format(mm_lsbbit, "000") & "]=" + _
                               m_bitStrM + "=" + m_HexStr + "=" + CStr(CFGTable.Category(i).Cate32bit(k).Decimal)
                    TheExec.Datalog.WriteComment Space(5) + m_tmpStr
                Next k
                TheExec.Datalog.WriteComment ""
                Exit For
            End If
        Next i
    End If

    auto_display_CFG_Cond_Table_by_PKGName = 1

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20171103 add
Public Function UDRE_Index(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "UDRE_Index"

    Dim i As Long
    Dim match_Flag As Boolean

    match_Flag = False
'    For i = 0 To UBound(UDRE_Fuse.Category)
'        If (UCase(myStr) = UCase(UDRE_Fuse.Category(i).name)) Then
'            UDRE_Index = i
'            match_Flag = True
'            Exit For
'        End If
'    Next i
    If (True) Then
        Dim m_keyname As String
        m_keyname = myStr
        UDRE_Index = eFuse_GetStoredIndex(eFuse_UDRE, m_keyname)
        If (UDRE_Index >= 0) Then match_Flag = True
    Else
    For i = 0 To UBound(UDRE_Fuse.Category)
        If (UCase(myStr) = UCase(UDRE_Fuse.Category(i).Name)) Then
            UDRE_Index = i
            match_Flag = True
            Exit For
        End If
    Next i
    End If

    If (match_Flag = False) Then
        UDRE_Index = -1
        PrintDataLog "UDRE_Index:: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
''''20171103 add
Public Function UDRP_Index(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "UDRP_Index"

    Dim i As Long
    Dim match_Flag As Boolean

    match_Flag = False
'    For i = 0 To UBound(UDRP_Fuse.Category)
'        If (UCase(myStr) = UCase(UDRP_Fuse.Category(i).name)) Then
'            UDRP_Index = i
'            match_Flag = True
'            Exit For
'        End If
'    Next i
    If (True) Then
        Dim m_keyname As String
        m_keyname = myStr
        UDRP_Index = eFuse_GetStoredIndex(eFuse_UDRP, m_keyname)
        If (UDRP_Index >= 0) Then match_Flag = True
    Else
    For i = 0 To UBound(UDRP_Fuse.Category)
        If (UCase(myStr) = UCase(UDRP_Fuse.Category(i).Name)) Then
            UDRP_Index = i
            match_Flag = True
            Exit For
        End If
    Next i
    End If

    If (match_Flag = False) Then
        UDRP_Index = -1
        PrintDataLog "UDRP_Index:: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20171103 add
Public Function CMPE_Index(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CMPE_Index"

    Dim i As Long
    Dim match_Flag As Boolean

    match_Flag = False
'    For i = 0 To UBound(CMPE_Fuse.Category)
'        If (UCase(myStr) = UCase(CMPE_Fuse.Category(i).name)) Then
'            CMPE_Index = i
'            match_Flag = True
'            Exit For
'        End If
'    Next i
    If (True) Then
        Dim m_keyname As String
        m_keyname = myStr
        CMPE_Index = eFuse_GetStoredIndex(eFuse_CMPE, m_keyname)
        If (CMPE_Index >= 0) Then match_Flag = True
    Else
    For i = 0 To UBound(CMPE_Fuse.Category)
        If (UCase(myStr) = UCase(CMPE_Fuse.Category(i).Name)) Then
            CMPE_Index = i
            match_Flag = True
            Exit For
        End If
    Next i
    End If

    If (match_Flag = False) Then
        CMPE_Index = -1
        PrintDataLog "CMPE_Index:: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20171103 add
Public Function CMPP_Index(myStr As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CMPP_Index"

    Dim i As Long
    Dim match_Flag As Boolean

    match_Flag = False
'    For i = 0 To UBound(CMPP_Fuse.Category)
'        If (UCase(myStr) = UCase(CMPP_Fuse.Category(i).name)) Then
'            CMPP_Index = i
'            match_Flag = True
'            Exit For
'        End If
'    Next i
    If (True) Then
        Dim m_keyname As String
        m_keyname = myStr
        CMPP_Index = eFuse_GetStoredIndex(eFuse_CMPP, m_keyname)
        If (CMPP_Index >= 0) Then match_Flag = True
    Else
    For i = 0 To UBound(CMPP_Fuse.Category)
        If (UCase(myStr) = UCase(CMPP_Fuse.Category(i).Name)) Then
            CMPP_Index = i
            match_Flag = True
            Exit For
        End If
    Next i
    End If


    If (match_Flag = False) Then
        CMPP_Index = -1
        PrintDataLog "CMPP_Index:: <" + myStr + ">, it's NOT existed in the Category."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20171211 update
Public Function auto_isBinaryString(ByVal InputStr As String) As Boolean

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_isBinaryString"
    
    Dim i As Long
    Dim m_len As Long
    Dim m_char As String
    Dim m_match_flag As Boolean

    InputStr = UCase(Trim(InputStr))

    If (InputStr Like "B*") Then
        m_match_flag = True ''''<MUST> initialize
        InputStr = Replace(InputStr, "B", "", 1, 1) ''''remove the first "B" character
        InputStr = Replace(InputStr, "_", "")       ''''for case like as b0001_0101

        ''''do the advanced analysis
        m_len = Len(InputStr)
        For i = 1 To m_len
            m_char = Mid(InputStr, i, 1)
            If (m_char <> "0" And m_char <> "1") Then
                m_match_flag = False
                Exit For
            End If
        Next i
    Else
        m_match_flag = False
    End If

    auto_isBinaryString = m_match_flag

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20171211 update
'Public Function auto_checkDefaultValue(ByVal m_defval As Variant, m_binarr() As Long, Optional m_bitwidth As Long = 0, _
'                                       Optional m_defreal As String = "NA") As Variant
Public Function auto_checkDefaultValue(ByVal m_defval As Variant, ByVal m_alogrithm As String, m_binarr() As Long, Optional m_bitwidth As Long = 0, _
                                       Optional m_defreal As String = "NA") As Variant


On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_checkDefaultValue"

    Dim m_defval_orig As Variant
    Dim m_defvalhex As String
    Dim m_HexStr As String
    Dim m_match_flag As Boolean
    Dim i As Long
    Dim m_len As Long
    Dim m_bitStrM As String

    m_defval_orig = m_defval
    m_match_flag = False

    ''''Here it's used to judge default value if is Hex or Binary
    If (auto_isHexString(CStr(m_defval))) Then
    
        ''''20180522 update
        m_defvalhex = Replace(UCase(CStr(m_defval)), "0X", "", 1, 1)
        m_defvalhex = Replace(UCase(CStr(m_defval)), "X", "", 1, 1)
        m_defvalhex = Replace(m_defvalhex, "_", "")
        
        ''''20160620 update
        If (auto_chkHexStr_isOver7FFFFFFF(CStr(m_defval)) = False) Then
            m_defval = CLng("&H" & m_defvalhex) ''''Here it's Hex2Dec
        Else
            ''''<MUST> keep the prefix "0x"
            ''''In function "auto_chkHexStr_isOver7FFFFFFF" will check allZero case as 0 only
        End If
        m_HexStr = Replace(CStr(m_defval_orig), "_", "")
        m_match_flag = True
    ElseIf (auto_isBinaryString(CStr(m_defval))) Then ''''20171211 add
        Dim m_BinStr As String
        m_BinStr = Replace(UCase(CStr(m_defval)), "B", "", 1, 1) ''''remove the first "B" character
        m_BinStr = Replace(m_BinStr, "_", "")                    ''''for case like as b0001_0101
        
        ''''convert to Hex with the prefix "0x"
        m_defvalhex = "0x" + auto_BinStr2HexStr(m_BinStr, 1)
        m_HexStr = m_defvalhex
        If (auto_chkHexStr_isOver7FFFFFFF(CStr(m_defvalhex)) = False) Then
            m_defvalhex = Replace(UCase(CStr(m_defvalhex)), "0X", "", 1, 1)
            m_defval = CLng("&H" & m_defvalhex) ''''Here it's Hex2Dec
        Else
            ''''with prefix 0x
            m_defval = m_defvalhex
        End If
        m_match_flag = True
    
    Else
        ''''value to HexStr
        m_HexStr = auto_Value2HexStr(CStr(m_defval), m_bitwidth)
    End If

    auto_checkDefaultValue = m_defval
    
    If (m_bitwidth > 0) Then ReDim m_binarr(m_bitwidth - 1) ''''<MUST>
    
    If (LCase(m_defreal) <> "real" And LCase(m_defreal) <> "bincut" And _
            Not ((LCase(m_alogrithm) = "vddbin" Or LCase(m_alogrithm) = "base") And LCase(m_defreal) = "default")) Then
    'If (LCase(m_defreal) <> "real" And LCase(m_defreal) <> "bincut" And Not (LCase(m_defreal) Like "safe*voltage")) Then
        m_match_flag = m_match_flag And True
        ''''20180711 New for binarr(), Here it MUST bypass "safe voltage"
        m_bitStrM = auto_HexStr2BinStr_EFUSE(m_HexStr, m_bitwidth, m_binarr)
    Else
        m_match_flag = False
    End If

    ''''check if it exceeds over the value of the bitwidth
    Dim m_tmpStr As String
    Dim m_value As Variant
    Dim m_upperValue As Double
    If (m_defval <> 0 And m_bitwidth > 0 And m_match_flag = True) Then
        ''''20180522 update
        m_value = auto_HexStr2Value(m_HexStr)
        If (m_bitwidth > 1023 And auto_isHexString(m_value) = True) Then
            TheExec.Datalog.WriteComment "<Check> " + funcName + ":: BitWidth(" + CStr(m_bitwidth) + ") is over 1023bits (CDbl limit), Vaule=" + CStr(m_value)
        Else
            m_upperValue = CDbl(2 ^ m_bitwidth - 1)
            If (m_value > m_upperValue) Then
                m_tmpStr = "<WARNING> " + funcName + ":: input value (" + CStr(m_defval_orig) + "=" + CStr(m_value)
                m_tmpStr = m_tmpStr + ") > BitWidth(" + CStr(m_bitwidth) + ") Limit Vaule (" + CStr(m_upperValue) + ")"
                TheExec.Datalog.WriteComment m_tmpStr
                TheExec.AddOutput m_tmpStr
                auto_checkDefaultValue = -999 ''''"error"
                GoTo errHandler
            End If
        End If
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20170126 New (20180122 add from Tosp)
''''It's used in serial 1-bit (JTAG_TDO) read of eFuse, then convert to 32bits binary string
Public Function auto_eFuse_DSSC_ReadDigCap_1bit_to_32bits(cycleNum As Long, ByRef SingleStrArray() As String, CapWave As DSPWave, ByRef blank As SiteBoolean, _
                                                          Optional PatBitOrder As String = "bit0_bitLast")

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_DSSC_ReadDigCap_1bit_to_32bits"
    
    Dim i As Long, k As Long
    Dim m_size As Long
    Dim N1 As Long
    Dim m_32bitsBinStr As String ''''[bit31...bit0]
    Dim site As Variant
    Dim m_dataArr() As Long
    ReDim SingleStrArray(cycleNum - 1, TheExec.sites.Existing.Count - 1)
    
    'initial
    blank = True
''    For Each Site In TheExec.Sites
''        blank(Site) = True
''    Next Site
    
    m_size = CapWave.SampleSize
    ReDim m_dataArr(m_size - 1)
    ''''build up array DAPWave.Element()
    If (UCase(PatBitOrder) = UCase("bit0_bitLast")) Then
        For Each site In TheExec.sites
            ''ReDim m_dataArr(m_size - 1) ''''reset and initialize
            m_dataArr = CapWave.Data

            ''''Here m_dataArr(0) is the bit0
            ''''     m_dataArr(m_size - 1) is the bitLast
            For N1 = 0 To cycleNum - 1
                m_32bitsBinStr = ""
                For i = 0 To 31 ''''compose a 32bits binary string
                    k = N1 * 32 + i
                    m_32bitsBinStr = CStr(m_dataArr(k)) + m_32bitsBinStr
                    If m_dataArr(k) <> 0 Then blank(site) = False
                Next i
                SingleStrArray(N1, site) = m_32bitsBinStr
            Next N1
        Next site
    Else
        ''''case "bitLast_bit0"
        For Each site In TheExec.sites
            ''ReDim m_dataArr(m_size - 1) ''''reset and initialize
            m_dataArr = CapWave.Data
            
            ''''Here m_dataArr(0) is the bitLast
            ''''     m_dataArr(m_size - 1) is the bit0
            For N1 = (cycleNum - 1) To 0 Step -1
                m_32bitsBinStr = ""
                For i = 0 To 31 ''''compose a 32bits binary string
                    k = ((cycleNum - 1) - N1) * 32 + i
                    m_32bitsBinStr = CStr(m_dataArr(k)) + m_32bitsBinStr
                    If m_dataArr(k) <> 0 Then blank(site) = False
                Next i
                SingleStrArray(N1, site) = m_32bitsBinStr
            Next N1
        Next site
    End If

    TheHdw.Wait 0.0001

Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function




''''20180522 update for the case with the same value on allSites
Public Function eFuse_DSSC_SetupDigSrcArr_allSites(patt As String, DigSrcPin As PinList, SignalName As String, SegmentSize As Long, WaveDefArray() As Long)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "eFuse_DSSC_SetupDigSrcArr_allSites"
    
    Dim InWave As New DSPWave
    Dim waveDblArray() As Double
    Dim site As Variant
    Dim WaveDef As String

    For Each site In TheExec.sites.Active
        InWave.Data = WaveDefArray
        InWave = InWave.ConvertDataTypeTo(DspDouble)
        waveDblArray = InWave.Data
        Exit For
    Next site
    
    WaveDef = "WaveDef_" + SignalName + "_allSites"
    TheHdw.Patterns(patt).Load
    
    ''''<NOTICE> Here WaveDefArray() must be Double for this case
    TheExec.WaveDefinitions.CreateWaveDefinition WaveDef, waveDblArray, True
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName
    
    ''''<NOTICE> check if there is already one outside
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
    
    With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
        .WaveDefinitionName = WaveDef
        .SampleSize = SegmentSize
        .Amplitude = 1
        '.LoadSamples
        .LoadSettings
    End With

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'''''20180522 update for the case with the same value on allSites
'Public Function eFuse_DSSC_SetupDigSrcWave_allSites(patt As String, digSrcPin As PinList, SignalName As String, SegmentSize As Long, srcWave As DSPWave)
'On Error GoTo errHandler
'    Dim funcName As String:: funcName = "eFuse_DSSC_SetupDigSrcWave_allSites"
'
'    Dim InWave As New DSPWave
'    Dim waveDblArray() As Double
'    Dim site As Variant
'    Dim WaveDef As String
'
'    For Each site In TheExec.sites.Active
'        InWave = srcWave.ConvertDataTypeTo(DspDouble)
'        waveDblArray = InWave.Data
'        Exit For
'    Next site
'
'    WaveDef = "WaveDef_" + SignalName + "_allSites"
'    TheHdw.Patterns(patt).Load
'
'    ''''<NOTICE> Here waveDblArray() must be Double for this case
'    TheExec.WaveDefinitions.CreateWaveDefinition WaveDef, waveDblArray, True
'    TheHdw.DSSC.Pins(digSrcPin).pattern(patt).Source.Signals.Add SignalName
'
'    ''''<NOTICE> check if there is already one outside
'    ''TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
'
'    With TheHdw.DSSC.Pins(digSrcPin).pattern(patt).Source.Signals(SignalName)
'        .WaveDefinitionName = WaveDef
'        .SampleSize = SegmentSize
'        .Amplitude = 1
'        .LoadSamples
'        .LoadSettings
'    End With
'
'Exit Function
'
'errHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function


Public Function DSSC_SetupDigSrcWave_TTR(patt As String, DigSrcPin As PinList, SignalName As String, SegmentSize As Long, InWave As DSPWave)
    'store efuse program bit into a DSP wave
    'Dim InWave As New DSPWave
    Dim site As Variant
    Dim WaveDef As String
    WaveDef = "WaveDef"
    'InWave.Data = WaveDefArray
    site = TheExec.sites.SiteNumber
    
    TheHdw.Patterns(patt).Load
    TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & site, InWave, True
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName
    With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
        .WaveDefinitionName = WaveDef & site
        .SampleSize = SegmentSize
        .Amplitude = 1
        .LoadSamples
        .LoadSettings
    End With
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
End Function

Public Function auto_eFuse_DSSC_ReadDigCap_Serial_1bits_to_32bitsPerRow(cycleNum As Long, PinName As String, ByRef SingleStrArray() As String, CapWave As DSPWave, ByRef blank As SiteBoolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_DSSC_ReadDigCap_Serial_1bits_to_32bitsPerRow"
    
    Dim Pin_Ary() As String, Pin_Cnt As Long, p_indx As Long, N1 As Long
    Dim ByteString As String, CurrInstance As String
    Dim hram_pindata As New PinListData
    Dim Cdata As String
    Dim p As Variant, p_idx As Long
    Dim AlarmStr As String
    Dim site As Variant
    Dim count32bit As Long
    ReDim SingleStrArray(cycleNum - 1, TheExec.sites.Existing.Count - 1)
    Dim idx As Long
    
    'initial
    For Each site In TheExec.sites
        blank(site) = True
    Next site
    
    For Each site In TheExec.sites
        For N1 = 0 To cycleNum - 1
            ByteString = ""
            'For count32bit = 0 To 31
            For count32bit = 31 To 0 Step -1
                idx = count32bit + N1 * 32
                ByteString = ByteString & CStr(CapWave.Element(idx))  ''auto_eFuse_Dec2BinStr32Bit(32, CapWave.Element(N1))
                If CapWave.Element(idx) <> 0 Then blank(site) = False    'return blank bolean for ecid blank check
            Next count32bit
            SingleStrArray(N1, site) = ByteString
        Next N1
    Next site

    TheHdw.Wait 0.001

Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function


'Public Function Get_BKM_File()
'
'On Error GoTo errHandler
'
'    Dim funcName As String:: funcName = "Get_BKM_File"
'    Dim m_file As String 'file location ex: C:\aaa.txt
'    Dim m_lineStr As String
'    Dim tok() As String
'    Dim split_content() As String
'    Set gO_BKM_Dic = CreateObject("Scripting.Dictionary")
'
'    'm_file = "D:\BKM\BKM_TABLE.txt"
'    m_file = "X:\Production\TMKF47\A0\BKM_TABLE.txt"
'
''1 BKM4
''2 BKM6.1
''3 BKM6.2
''4 BKM6.2M
''5 BKM6.3
''6 BKM6.4
'
'    If (Dir(m_file) = "") Then
'        TheExec.Datalog.WriteComment "<Error> BKM File:: " + m_file + " is NOT existed, please check it out!!"
'        Exit Function ''''MUST have
'    End If
'
'    Open m_file For Input As #1
'        Do Until EOF(1)
'            Line Input #1, m_lineStr
'                m_lineStr = Trim(m_lineStr)
'
'            tok = Split(m_lineStr, " ") 'ex: 2 BKM6.1
'            'tok(0) '2
'            'tok(1) 'BKM6.1
'
'            split_content = Split(tok(1), "BKM")
'            'split_conten(1)===>6.1
'
'            If (UCase(gS_JobName) Like "CP*") Then
'                If Not (gO_BKM_Dic.Exists(split_content(1))) Then
'                    gO_BKM_Dic.Add split_content(1), CDbl(tok(0))   '''6.1, 2
'                    'Add the new string with the match number into the dictionary
'                End If
'            ElseIf (UCase(gS_JobName) Like "*FT*") Then
'                If Not (gO_BKM_Dic.Exists(tok(0))) Then
'                    gO_BKM_Dic.Add tok(0), split_content(1)  '''2, 6.1
'                    'Add the new string with the match number into the dictionary
'                End If
'            Else
'                                TheExec.Datalog.WriteComment "<Error> " + funcName + ":: has no define job, please check it out."
'            End If
'        Loop
'    Close #1
'
'
'Exit Function
'
'errHandler:
'    Close #1
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function

''''20180522 update for the case with the same value on allSites
Public Function eFuse_DSSC_SetupDigSrcWave_allSites(patt As String, DigSrcPin As PinList, SignalName As String, srcWave As DSPWave)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "eFuse_DSSC_SetupDigSrcWave_allSites"
    
    Dim InWave As New DSPWave
    Dim waveDblArray() As Double
    Dim site As Variant
    Dim WaveDef As String
    Dim m_segsize As Long

    For Each site In TheExec.sites.Active
        InWave = srcWave.ConvertDataTypeTo(DspDouble).Copy
        waveDblArray = InWave.Data
        m_segsize = InWave.SampleSize
        Exit For
    Next site
    
    WaveDef = "WaveDef_" + SignalName + "_allSites"
    TheHdw.Patterns(patt).Load
    
    ''''<NOTICE> Here waveDblArray() must be Double for this case
    TheExec.WaveDefinitions.CreateWaveDefinition WaveDef, waveDblArray, True
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName

    With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
        .WaveDefinitionName = WaveDef
        .SampleSize = m_segsize
        .Amplitude = 1
        ''.LoadSamples ''''could waste TT and break PTE
        .LoadSettings
    End With

    ''''<NOTICE> check if there is already one outside
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20180831 New
''''return: 1  => Job > Stage
''''return: 0  => Job = Stage
''''return: -1 => Job < Stage
Public Function auto_eFuse_check_Job_cmpare_Stage(stageName As String, Optional showPrint As Boolean = "False") As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_check_Job_cmpare_Stage"

    Dim m_jobNum As Long
    Dim m_stageNum As Long

    ''''----------------------------------------------------------
    ''''standard production flow is
    '''' CP1_EARLY-->CP1-->CP2-->CP3-->WLFT-->FT1(FT1_25C)-->FT2(FT2_85C)-->FT3 (-->FT4-->FT5)
    ''''----------------------------------------------------------

    Select Case UCase(gS_JobName)
        Case "CP1_EARLY" ''''20171016 update
            m_jobNum = 0
        Case "CP1"
            m_jobNum = 1
        Case "CP2"
            m_jobNum = 2
        Case "CP3"
            m_jobNum = 3
        Case "WLFT", "WLFT1"
            m_jobNum = 9
        Case "FT1_EARLY"
            m_jobNum = 10
        Case "FT1"
            m_jobNum = 11
        Case "FT1_25C"
            m_jobNum = 12
        Case "FT2"
            m_jobNum = 13
        Case "FT2_85C"
            m_jobNum = 14
        Case "FT3"
            m_jobNum = 15
        Case "FT4"
            m_jobNum = 16
        Case "FT5"
            m_jobNum = 17
        Case Else
            m_jobNum = 99
    End Select
   
    Select Case UCase(stageName)
        Case "CP1_EARLY" ''''20171016 update
            m_stageNum = 0
        Case "CP1"
            m_stageNum = 1
        Case "CP2"
            m_stageNum = 2
        Case "CP3"
            m_stageNum = 3
        Case "WLFT", "WLFT1"
            m_stageNum = 9
        Case "FT1_EARLY"
            m_stageNum = 10
        Case "FT1"
            m_stageNum = 11
        Case "FT1_25C"
            m_stageNum = 12
        Case "FT2"
            m_stageNum = 13
        Case "FT2_85C"
            m_stageNum = 14
        Case "FT3"
            m_stageNum = 15
        Case "FT4"
            m_stageNum = 16
        Case "FT5"
            m_stageNum = 17
        Case Else
            m_stageNum = 99
    End Select

    If (m_jobNum > m_stageNum) Then
        ''''means that (Job > Stage)
        ''''means that setLimit = 0
        auto_eFuse_check_Job_cmpare_Stage = 1
    ElseIf (m_jobNum = m_stageNum) Then
        ''''means that (Job = Stage)
        ''''means that setLimit as sheet
        auto_eFuse_check_Job_cmpare_Stage = 0
        showPrint = False
    ElseIf (m_jobNum < m_stageNum) Then
        ''''means that (Job < Stage)
        ''''means that setLimit as sheet
        auto_eFuse_check_Job_cmpare_Stage = -1
    End If

    If (showPrint = True) Then
        Dim m_tmpStr As String
        m_tmpStr = funcName + "=" + CStr(auto_eFuse_check_Job_cmpare_Stage)
        m_tmpStr = m_tmpStr + " :: Job = " + UCase(gS_JobName) + "(m_jobNum=" + CStr(m_jobNum) + ")"
        m_tmpStr = m_tmpStr + ", Pgm_Stage = " + UCase(stageName) + "(m_stageNum=" + CStr(m_stageNum) + ")"
        TheExec.Datalog.WriteComment m_tmpStr
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''201808XX New
Public Function auto_eFuse_param2globalDSPVar(fuseblock As eFuseBlockType) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_param2globalDSPVar"

    Dim i As Long
    Dim j As Long
    Dim m_Site As Variant
    Dim BitsPerRow As Long
    Dim ReadCycles As Long
    Dim BitsPerCycle As Long
    Dim BitsPerBlock As Long
    Dim DigSrcRepeatN As Long

    If (fuseblock = eFuse_ECID) Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = EcidBitsPerRow             ''''32   , 16  , 32
        ReadCycles = EcidReadCycle              ''''16   , 16  , 16
        BitsPerCycle = ECIDBitPerCycle          ''''32   , 32  , 32
        BitsPerBlock = EcidBitPerBlockUsed      ''''256  , 256 , 512
        DigSrcRepeatN = EcidWriteBitExpandWidth

    ElseIf (fuseblock = eFuse_CFG) Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = EConfigBitsPerRow          ''''32   , 16  , 32
        ReadCycles = EConfigReadCycle           ''''32   , 32  , 16
        BitsPerCycle = EConfigReadBitWidth      ''''32   , 32  , 32
        BitsPerBlock = EConfigBitPerBlockUsed   ''''512  , 512 , 512
        DigSrcRepeatN = EConfig_Repeat_Cyc_for_Pgm

    ElseIf (fuseblock = eFuse_UID) Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = UIDBitsPerRow              ''''32   , 16  , 32
        ReadCycles = UIDReadCycle               ''''64   , 64  , 32
        BitsPerCycle = UIDBitsPerCycle          ''''32   , 32  , 32
        BitsPerBlock = UIDBitsPerBlockUsed      ''''1024 , 1024, 1024
        DigSrcRepeatN = UIDWriteBitExpandWidth

    ElseIf (fuseblock = eFuse_SEN) Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = SENSORBitsPerRow           ''''32   , 16  , 32
        ReadCycles = SENSORReadCycle            ''''32   , 32  , 16
        BitsPerCycle = SENSORReadBitWidth       ''''32   , 32  , 32
        BitsPerBlock = SENSORBitPerBlockUsed    ''''512  , 512 , 512
        DigSrcRepeatN = SENSOR_Repeat_Cyc_for_Pgm

    ElseIf (fuseblock = eFuse_MON) Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = MONITORBitsPerRow          ''''32   , 16  , 32
        ReadCycles = MONITORReadCycle           ''''32   , 32  , 16
        BitsPerCycle = MONITORReadBitWidth      ''''32   , 32  , 32
        BitsPerBlock = MONITORBitPerBlockUsed   ''''512  , 512 , 512
        DigSrcRepeatN = MONITOR_Repeat_Cyc_for_Pgm
        
    ElseIf (fuseblock = eFuse_UDR) Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = 1                          ''''32   , 16  , 32
        ReadCycles = gL_USI_DigSrcBits_Num           ''''32   , 32  , 16
        BitsPerCycle = 1                        ''''32   , 32  , 32
        BitsPerBlock = gL_USI_DigSrcBits_Num    ''''512  , 512 , 512
        DigSrcRepeatN = 1
        
    ElseIf (fuseblock = eFuse_UDRE) Then
                                                    ''''U2D  , R2L , SUP
        BitsPerRow = 1                              ''''32   , 16  , 32
        ReadCycles = gL_UDRE_USI_DigSrcBits_Num     ''''32   , 32  , 16
        BitsPerCycle = 1                            ''''32   , 32  , 32
        BitsPerBlock = gL_UDRE_USI_DigSrcBits_Num   ''''512  , 512 , 512
        DigSrcRepeatN = 1
        
    ElseIf (fuseblock = eFuse_UDRP) Then
                                                    ''''U2D  , R2L , SUP
        BitsPerRow = 1                              ''''32   , 16  , 32
        ReadCycles = gL_UDRP_USI_DigSrcBits_Num     ''''32   , 32  , 16
        BitsPerCycle = 1                            ''''32   , 32  , 32
        BitsPerBlock = gL_UDRP_USI_DigSrcBits_Num   ''''512  , 512 , 512
        DigSrcRepeatN = 1
    ElseIf (fuseblock = eFuse_CMP) Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = 1                          ''''32   , 16  , 32
        ReadCycles = gL_CMP_DigCapBits_Num           ''''32   , 32  , 16
        BitsPerCycle = 1                        ''''32   , 32  , 32
        BitsPerBlock = gL_CMP_DigCapBits_Num    ''''512  , 512 , 512
        DigSrcRepeatN = 1

    ElseIf (fuseblock = eFuse_CMPE) Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = 1                          ''''32   , 16  , 32
        ReadCycles = gL_CMPE_DigCapBits_Num           ''''32   , 32  , 16
        BitsPerCycle = 1                        ''''32   , 32  , 32
        BitsPerBlock = gL_CMPE_DigCapBits_Num    ''''512  , 512 , 512
        DigSrcRepeatN = 1
        
    ElseIf (fuseblock = eFuse_CMPP) Then
                                                ''''U2D  , R2L , SUP
        BitsPerRow = 1                          ''''32   , 16  , 32
        ReadCycles = gL_CMPP_DigCapBits_Num           ''''32   , 32  , 16
        BitsPerCycle = 1                        ''''32   , 32  , 32
        BitsPerBlock = gL_CMPP_DigCapBits_Num    ''''512  , 512 , 512
        DigSrcRepeatN = 1
        
    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,SEN,MON)"
        GoTo errHandler
        ''''nothing
    End If

    gDL_BitsPerRow = BitsPerRow
    gDL_ReadCycles = ReadCycles
    gDL_BitsPerCycle = BitsPerCycle
    gDL_BitsPerBlock = BitsPerBlock
    gDL_TotalBits = BitsPerCycle * ReadCycles
    gDL_DigSrcRepeatN = DigSrcRepeatN
    gDD_BaseVoltage = gD_BaseVoltage
    gDD_BaseStepVoltage = gD_BaseStepVoltage
    gDB_SerialType = False
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''201808XX new method with DspWave
''''201812XX Update new method with global DspWave
''Public Function auto_eFuse_DSSC_ReadDigCap_32bits_NEW(fuseblock As eFuseBlockType, capWave As DSPWave, bitFlagWave As DSPWave, _
''                                                      ByRef singleBitWave As DSPWave, ByRef doubleBitWave As DSPWave, ByRef FBCount As SiteLong, _
''                                                      ByRef blank_stage As SiteBoolean, ByRef allblank As SiteBoolean, _
''                                                      Optional serialCap As Boolean = False, Optional PatBitOrder As String = "bit0_bitLast")
''''201812XX update method
Public Function auto_eFuse_DSSC_ReadDigCap_32bits_NEW(fuseblock As eFuseBlockType, ByVal bitFlag_mode As Long, ByVal CapWave As DSPWave, _
                                                      ByRef FBCount As SiteLong, ByRef blank_stage As SiteBoolean, ByRef allBlank As SiteBoolean, _
                                                      Optional serialCap As Boolean = False, Optional PatBitOrder As String = "bit0_bitLast")

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_DSSC_ReadDigCap_32bits_NEW"
    
    Dim m_Site As Variant
    Dim cycleNum As Long
    Dim m_dlogstr As String
    Dim m_blank As New SiteBoolean
    Dim mSL_simBlank As New SiteLong
    Dim m_jobname_pre As String
    Dim m_caseFlag As New SiteBoolean

    ''''move to outside
    ''Call auto_eFuse_param2globalDSPVar(fuseblock) ''''<MUST be here First>
    
    ''''After the above call, the cycleNum can be got.
    cycleNum = gDL_ReadCycles

    ''''--------------------------------------------------------------------------
    '''' Simulation Start                                                        |
    ''''--------------------------------------------------------------------------
'If (fuseblock = eFuse_CFG Or fuseblock = eFuse_ECID Or fuseblock = eFuse_MON Or fuseblock = eFuse_SEN) Then
If (fuseblock <> eFuse_UID) Then
If (TheExec.TesterMode = testModeOffline) Then
        ''''try
        'serialCap = True ''True or False
        'PatBitOrder = "bitLast_bit0" ''"bit0_bitLast"

        If (gL_eFuse_Sim_Blank = 0) Then
            m_blank = True
        ElseIf (gL_eFuse_Sim_Blank = 1) Then
            m_blank = False
        Else
            ''''can be used to try the different scenario on the different site
            Dim m_seq As Long
            For Each m_Site In TheExec.sites
                If (m_seq = 0) Then
                    m_blank = True
                    m_seq = 1
                Else
                    m_blank = False
                    m_seq = 0
                End If
            Next m_Site
        End If

        m_jobname_pre = gS_JobName

        ''''<Trick> 201812XX update
        ''''Here using Site loop can simulate the different scenario on the different site
        ''''------------------------------------------------------------------------------
        ''''  simBlank: simulate blank condition to decide which stage bit flag to be used
        ''''       = 0: means that all bits blank=True as early stage bits
        ''''       = 1: means that simulate those bits (stage <  job)
        ''''       = 2: means that simulate those bits (stage <= job)
        ''''------------------------------------------------------------------------------
        TheExec.Datalog.WriteComment ""
        For Each m_Site In TheExec.sites
            m_dlogstr = funcName + ":: Site(" & m_Site & ") Blank = " & m_blank(m_Site)
            If (m_blank = True And gS_JobName = "cp1_early") Then
                mSL_simBlank = 0
                m_dlogstr = m_dlogstr + " (Simulate Bits Stage = " + gS_JobName + ")"
            ElseIf (m_blank = True And gS_JobName <> "cp1_early") Then
                mSL_simBlank = 1
                m_dlogstr = m_dlogstr + " (Simulate Bits Stage < " + gS_JobName + ")"
            ElseIf (m_blank = False) Then ''''if blank=False, Retest mode
                mSL_simBlank = 2
                ''''<Importance> simulate cp1+cp1_early when the blank=false
                If (m_jobname_pre = "cp1_early") Then gS_JobName = "cp1"
                m_dlogstr = m_dlogstr + " (Simulate Bits Stage <= " + gS_JobName + ")"
            End If
            TheExec.Datalog.WriteComment m_dlogstr
            If (PatBitOrder = "bit0_bitLast") Then
                m_caseFlag = False
            ElseIf (PatBitOrder = "bitLast_bit0") Then
                m_caseFlag = True
            End If
            
            Call eFuseENGFakeValue_Sim
            
            'Call auto_eFuse_Simulate_fromWrite2CapWave(fuseblock, mSL_simBlank, capWave)
            Call auto_eFuse_Simulate_fromWrite2CapWave(fuseblock, mSL_simBlank, CapWave, m_caseFlag)
        Next m_Site

        ''''<NOTICE and MUST> reset for the case "mSL_simBlank = 2"
        If (m_blank.Any(False) And m_jobname_pre = "cp1_early") Then
            gS_JobName = m_jobname_pre
        End If
End If
End If
'End If
    ''''--------------------------------------------------------------------------
    '''' Simulation End                                                          |
    ''''--------------------------------------------------------------------------

    If (serialCap = False) Then
        ''''<TRICK and NOTICE>
        ''''Here using Site loop to make sure capWave is Ready when it's Automatic Mode
        For Each m_Site In TheExec.sites
            If (CapWave.SampleSize <> cycleNum) Then GoTo errHandler
        Next m_Site

        Call rundsp.eFuse_Wave32bits_to_SingleDoubleBitWave(fuseblock, bitFlag_mode, CapWave, FBCount, blank_stage, allBlank)
    
        ''''use for the debug purpose:: print the original capWave
        Call auto_eFuse_print_capWave32Bits(fuseblock, CapWave, False) ''''default to False, True: print dlog
    Else
        ''''serial capture
        'Dim m_caseFlag As New SiteBoolean
'        If (PatBitOrder = "bit0_bitLast") Then
'            m_caseFlag = True
'        ElseIf (PatBitOrder = "bitLast_bit0") Then
'            m_caseFlag = False
'        End If
        If (PatBitOrder = "bit0_bitLast") Then
            m_caseFlag = False
        ElseIf (PatBitOrder = "bitLast_bit0") Then
            m_caseFlag = True
        End If
        ''''update it later
        Call rundsp.eFuse_Wave1bit_to_SingleDoubleBitWave(fuseblock, bitFlag_mode, m_caseFlag, CapWave, FBCount, blank_stage, allBlank)
    End If
    TheHdw.Wait 1# * ms
    
Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

''''201811XX
Public Function auto_eFuse_ECID_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_ECID_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    m_arrdimSize = EcidBitPerBlockUsed - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(ECIDFuse.Category) - 1 ''''skip the last one "ECID_DEID" category

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)

    For i = 0 To m_size ''''skip last one
        With ECIDFuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
            
            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        ''''201812XX, could be used in the simulation mode
        If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then  '''' Job >= Stage
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
            Next j
        End If
    Next i

    ReDim gL_ECID_msbbit_arr(m_size)
    ReDim gL_ECID_lsbbit_arr(m_size)
    ReDim gL_ECID_bitwidth_arr(m_size)
    ReDim gL_ECID_DefaultOrReal_arr(m_size)
    ReDim gL_ECID_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_ECID_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_ECID_allDefaultBits_arr(m_arrdimSize)
    ReDim gL_ECID_stageLEQjob_bitFlag_arr(m_arrdimSize)

    gL_ECID_msbbit_arr = m_msbbit_arr
    gL_ECID_lsbbit_arr = m_lsbbit_arr
    gL_ECID_bitwidth_arr = m_bitwidth_arr
    gL_ECID_DefaultOrReal_arr = m_DefaultOrReal_arr
    
    gL_ECID_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_ECID_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_ECID_allDefaultBits_arr = m_allDefaultBits_arr
    gL_ECID_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''201811XX
Public Function auto_eFuse_CFG_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_CFG_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_stgae_real_bitFlag_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    ''''201812XX
    ''''-------------------------------
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_tmpStrM As String
    Dim m_bitwidth As Long
    Dim m_lolmt_binarr() As Long
    Dim m_hilmt_binarr() As Long
    Dim m_allLoLMTBits_arr() As Long
    Dim m_allHiLMTBits_arr() As Long
    Dim m_catenameVbin As String
    Dim MaxLevelIndex As Long
    Dim m_Pmode As Long
    ''''-------------------------------
    
    m_arrdimSize = EConfigBitPerBlockUsed - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(CFGFuse.Category)

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_stgae_real_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)

    ReDim m_allLoLMTBits_arr(m_arrdimSize)
    ReDim m_allHiLMTBits_arr(m_arrdimSize)

    For i = 0 To m_size
        With CFGFuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth

            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_resolution = .Resoultion
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            ''''including "decimal", "default"
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
                If ((m_dfreal = "real" And m_algorithm <> "cond") Or m_dfreal = "bincut") Then
                    For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                        m_stgae_real_bitFlag_arr(j) = 1
                    Next j
                End If
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
                If ((m_dfreal = "real" And m_algorithm <> "cond") Or m_dfreal = "bincut") Then
                    For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                        m_stgae_real_bitFlag_arr(j) = 1
                    Next j
                End If
            End If
        End If

        ''''201812XX, could be used in the simulation mode
        If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then  '''' Job >= Stage
            If (m_msbbit_arr(i) < m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''201812XX update
        ''''<MUST> Need to take care of Vddbin/Base/Bincut/IDS limit process
        ''''limit to bit array
''        m_bitwidth = m_bitwidth_arr(i)
''        If (m_algorithm = "base" Or m_algorithm = "vddbin") Then
''            If (m_DefaultOrReal_arr(i) = 1) Then
''                If (m_dfreal Like "safe*voltage") Then
''                    m_lolmt = CLng(m_lolmt / m_resolution)
''                    m_hilmt = m_lolmt
''                End If
''            Else
''                ''''it should NOT be 'real'
''            End If
''        ElseIf (m_algorithm = "vddbin") Then
''            If (m_DefaultOrReal_arr(i) = 1) Then
''                If (m_dfreal Like "safe*voltage") Then
''                    m_lolmt = (m_lolmt - gD_BaseVoltage) / m_resolution
''                    m_hilmt = m_lolmt
''                End If
''            Else
''                ''''it needs a siteloop [NOTICE]
''                ''''bincut
''                m_catenameVbin = CFGFuse.Category(i).Name
''                m_Pmode = VddBinStr2Enum(m_catenameVbin)
''                MaxLevelIndex = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step
''                m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(MaxLevelIndex) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(MaxLevelIndex)
''                m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
''                m_lolmt_binarr = m_DefValBitArr
''                m_hilmt_binarr = m_DefValBitArr
''            End If
''        ElseIf (m_algorithm = "ids") Then
''            If (CDbl(m_lolmt) = 0# And CDbl(m_hilmt) = 0#) Then ''''Need to check
''                m_lolmt = 1#
''                m_hilmt = (2 ^ m_bitwidth) - 1
''            Else
''                m_lolmt = CLng(m_lolmt / m_resolution)
''                If (m_lolmt = 0) Then m_lolmt = 1  '0 means nothing, can not be acceptable
''
''                m_hilmt = CLng(m_hilmt / m_resolution)
''                If (m_hilmt = (2 ^ m_bitwidth)) Then m_hilmt = (2 ^ m_bitwidth) - 1
''            End If
''            m_tmpStrM = auto_HexStr2BinStr_EFUSE(auto_Value2HexStr(CStr(m_lolmt)), m_bitwidth, m_lolmt_binarr)
''            m_tmpStrM = auto_HexStr2BinStr_EFUSE(auto_Value2HexStr(CStr(m_hilmt)), m_bitwidth, m_hilmt_binarr)
''        Else
''            ''''most case
''            m_tmpStrM = auto_HexStr2BinStr_EFUSE(auto_Value2HexStr(CStr(m_lolmt)), m_bitwidth, m_lolmt_binarr)
''            m_tmpStrM = auto_HexStr2BinStr_EFUSE(auto_Value2HexStr(CStr(m_hilmt)), m_bitwidth, m_hilmt_binarr)
''        End If


        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        End If
    Next i

    ReDim gL_CFG_msbbit_arr(m_size)
    ReDim gL_CFG_lsbbit_arr(m_size)
    ReDim gL_CFG_bitwidth_arr(m_size)
    ReDim gL_CFG_DefaultOrReal_arr(m_size)
    ReDim gL_CFG_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_CFG_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_CFG_stage_real_bitFlag_arr(m_arrdimSize)
    ReDim gL_CFG_allDefaultBits_arr(m_arrdimSize)
    ReDim gL_CFG_stageLEQjob_bitFlag_arr(m_arrdimSize)

    gL_CFG_msbbit_arr = m_msbbit_arr
    gL_CFG_lsbbit_arr = m_lsbbit_arr
    gL_CFG_bitwidth_arr = m_bitwidth_arr
    gL_CFG_DefaultOrReal_arr = m_DefaultOrReal_arr
    
    gL_CFG_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_CFG_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_CFG_stage_real_bitFlag_arr = m_stgae_real_bitFlag_arr
    gL_CFG_allDefaultBits_arr = m_allDefaultBits_arr
    gL_CFG_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''201904 Ter
Public Function auto_eFuse_MON_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse__MON_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    
    ''''201812XX
    ''''-------------------------------
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_tmpStrM As String
    Dim m_bitwidth As Long
    Dim m_lolmt_binarr() As Long
    Dim m_hilmt_binarr() As Long
    Dim m_allLoLMTBits_arr() As Long
    Dim m_allHiLMTBits_arr() As Long
    Dim m_catenameVbin As String
    Dim MaxLevelIndex As Long
    Dim m_Pmode As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    m_arrdimSize = MONITORBitPerBlockUsed - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(MONFuse.Category)

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    
    ReDim m_allLoLMTBits_arr(m_arrdimSize)
    ReDim m_allHiLMTBits_arr(m_arrdimSize)
    
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)

    For i = 0 To m_size
        With MONFuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth

            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_resolution = .Resoultion
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            ''''including "decimal", "default"
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then  '''' Job >= Stage
            If (m_msbbit_arr(i) < m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        End If
    Next i

    ReDim gL_MON_msbbit_arr(m_size)
    ReDim gL_MON_lsbbit_arr(m_size)
    ReDim gL_MON_bitwidth_arr(m_size)
    ReDim gL_MON_DefaultOrReal_arr(m_size)
    ReDim gL_MON_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_MON_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_MON_allDefaultBits_arr(m_arrdimSize)

    gL_MON_msbbit_arr = m_msbbit_arr
    gL_MON_lsbbit_arr = m_lsbbit_arr
    gL_MON_bitwidth_arr = m_bitwidth_arr
    gL_MON_DefaultOrReal_arr = m_DefaultOrReal_arr
    gL_MON_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_MON_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_MON_allDefaultBits_arr = m_allDefaultBits_arr
    gL_MON_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr


Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''201904 Ter
Public Function auto_eFuse_SEN_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_SEN_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    
    ''''201812XX
    ''''-------------------------------
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_tmpStrM As String
    Dim m_bitwidth As Long
    Dim m_lolmt_binarr() As Long
    Dim m_hilmt_binarr() As Long
    Dim m_allLoLMTBits_arr() As Long
    Dim m_allHiLMTBits_arr() As Long
    Dim m_catenameVbin As String
    Dim MaxLevelIndex As Long
    Dim m_Pmode As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    m_arrdimSize = SENSORBitPerBlockUsed - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(SENFuse.Category)

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    
    ReDim m_allLoLMTBits_arr(m_arrdimSize)
    ReDim m_allHiLMTBits_arr(m_arrdimSize)
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)

    For i = 0 To m_size
        With SENFuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth

            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_resolution = .Resoultion
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            ''''including "decimal", "default"
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then  '''' Job >= Stage
            If (m_msbbit_arr(i) < m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        End If
    Next i

    ReDim gL_SEN_msbbit_arr(m_size)
    ReDim gL_SEN_lsbbit_arr(m_size)
    ReDim gL_SEN_bitwidth_arr(m_size)
    ReDim gL_SEN_DefaultOrReal_arr(m_size)
    ReDim gL_SEN_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_SEN_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_SEN_allDefaultBits_arr(m_arrdimSize)

    gL_SEN_msbbit_arr = m_msbbit_arr
    gL_SEN_lsbbit_arr = m_lsbbit_arr
    gL_SEN_bitwidth_arr = m_bitwidth_arr
    gL_SEN_DefaultOrReal_arr = m_DefaultOrReal_arr
    gL_SEN_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_SEN_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_SEN_allDefaultBits_arr = m_allDefaultBits_arr
    gL_SEN_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''201904 Ter
Public Function auto_eFuse_UDR_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_UDR_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    
    ''''201812XX
    ''''-------------------------------
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_tmpStrM As String
    Dim m_bitwidth As Long
    Dim m_lolmt_binarr() As Long
    Dim m_hilmt_binarr() As Long
    Dim m_allLoLMTBits_arr() As Long
    Dim m_allHiLMTBits_arr() As Long
    Dim m_catenameVbin As String
    Dim MaxLevelIndex As Long
    Dim m_Pmode As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    m_arrdimSize = gL_USI_DigSrcBits_Num - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(UDRFuse.Category)

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    
    ReDim m_allLoLMTBits_arr(m_arrdimSize)
    ReDim m_allHiLMTBits_arr(m_arrdimSize)
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)

    For i = 0 To m_size
        With UDRFuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth

            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_resolution = .Resoultion
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            ''''including "decimal", "default"
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then  '''' Job >= Stage
            If (m_msbbit_arr(i) < m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        End If
    Next i

    ReDim gL_UDR_msbbit_arr(m_size)
    ReDim gL_UDR_lsbbit_arr(m_size)
    ReDim gL_UDR_bitwidth_arr(m_size)
    ReDim gL_UDR_DefaultOrReal_arr(m_size)
    ReDim gL_UDR_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_UDR_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_UDR_allDefaultBits_arr(m_arrdimSize)

    gL_UDR_msbbit_arr = m_msbbit_arr
    gL_UDR_lsbbit_arr = m_lsbbit_arr
    gL_UDR_bitwidth_arr = m_bitwidth_arr
    gL_UDR_DefaultOrReal_arr = m_DefaultOrReal_arr
    gL_UDR_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_UDR_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_UDR_allDefaultBits_arr = m_allDefaultBits_arr
    gL_UDR_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''201904 Ter
Public Function auto_eFuse_UID_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse__UID_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    
    ''''201812XX
    ''''-------------------------------
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_tmpStrM As String
    Dim m_bitwidth As Long
    Dim m_lolmt_binarr() As Long
    Dim m_hilmt_binarr() As Long
    Dim m_allLoLMTBits_arr() As Long
    Dim m_allHiLMTBits_arr() As Long
    Dim m_catenameVbin As String
    Dim MaxLevelIndex As Long
    Dim m_Pmode As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    m_arrdimSize = UIDBitsPerBlockUsed - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(UIDFuse.Category)

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    
    ReDim m_allLoLMTBits_arr(m_arrdimSize)
    ReDim m_allHiLMTBits_arr(m_arrdimSize)
    
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)
    
    For i = 0 To m_size
        With UIDFuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth

            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_resolution = .Resoultion
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            ''''including "decimal", "default"
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then  '''' Job >= Stage
            If (m_msbbit_arr(i) < m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        End If
    Next i

    ReDim gL_UID_msbbit_arr(m_size)
    ReDim gL_UID_lsbbit_arr(m_size)
    ReDim gL_UID_bitwidth_arr(m_size)
    ReDim gL_UID_DefaultOrReal_arr(m_size)
    ReDim gL_UID_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_UID_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_UID_allDefaultBits_arr(m_arrdimSize)

    gL_UID_msbbit_arr = m_msbbit_arr
    gL_UID_lsbbit_arr = m_lsbbit_arr
    gL_UID_bitwidth_arr = m_bitwidth_arr
    gL_UID_DefaultOrReal_arr = m_DefaultOrReal_arr
    gL_UID_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_UID_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_UID_allDefaultBits_arr = m_allDefaultBits_arr
    gL_UID_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''201904 Ter
Public Function auto_eFuse_UDRP_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_UDRP_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    
    ''''201812XX
    ''''-------------------------------
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_tmpStrM As String
    Dim m_bitwidth As Long
    Dim m_lolmt_binarr() As Long
    Dim m_hilmt_binarr() As Long
    Dim m_allLoLMTBits_arr() As Long
    Dim m_allHiLMTBits_arr() As Long
    Dim m_catenameVbin As String
    Dim MaxLevelIndex As Long
    Dim m_Pmode As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    m_arrdimSize = gL_UDRP_USI_DigSrcBits_Num - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(UDRP_Fuse.Category)

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    
    ReDim m_allLoLMTBits_arr(m_arrdimSize)
    ReDim m_allHiLMTBits_arr(m_arrdimSize)
    
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)

    For i = 0 To m_size
        With UDRP_Fuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth

            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_resolution = .Resoultion
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            ''''including "decimal", "default"
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then  '''' Job >= Stage
            If (m_msbbit_arr(i) < m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        End If
    Next i

    ReDim gL_UDRP_msbbit_arr(m_size)
    ReDim gL_UDRP_lsbbit_arr(m_size)
    ReDim gL_UDRP_bitwidth_arr(m_size)
    ReDim gL_UDRP_DefaultOrReal_arr(m_size)
    ReDim gL_UDRP_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_UDRP_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_UDRP_allDefaultBits_arr(m_arrdimSize)

    gL_UDRP_msbbit_arr = m_msbbit_arr
    gL_UDRP_lsbbit_arr = m_lsbbit_arr
    gL_UDRP_bitwidth_arr = m_bitwidth_arr
    gL_UDRP_DefaultOrReal_arr = m_DefaultOrReal_arr
    gL_UDRP_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_UDRP_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_UDRP_allDefaultBits_arr = m_allDefaultBits_arr
    gL_UDRP_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


''201904 Ter
Public Function auto_eFuse_UDRE_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_UDRE_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    
    ''''201812XX
    ''''-------------------------------
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_tmpStrM As String
    Dim m_bitwidth As Long
    Dim m_lolmt_binarr() As Long
    Dim m_hilmt_binarr() As Long
    Dim m_allLoLMTBits_arr() As Long
    Dim m_allHiLMTBits_arr() As Long
    Dim m_catenameVbin As String
    Dim MaxLevelIndex As Long
    Dim m_Pmode As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    m_arrdimSize = gL_UDRE_USI_DigSrcBits_Num - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(UDRE_Fuse.Category)

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    
    ReDim m_allLoLMTBits_arr(m_arrdimSize)
    ReDim m_allHiLMTBits_arr(m_arrdimSize)
    
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)

    For i = 0 To m_size
        With UDRE_Fuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth

            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_resolution = .Resoultion
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            ''''including "decimal", "default"
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then  '''' Job >= Stage
            If (m_msbbit_arr(i) < m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        End If
    Next i

    ReDim gL_UDRE_msbbit_arr(m_size)
    ReDim gL_UDRE_lsbbit_arr(m_size)
    ReDim gL_UDRE_bitwidth_arr(m_size)
    ReDim gL_UDRE_DefaultOrReal_arr(m_size)
    ReDim gL_UDRE_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_UDRE_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_UDRE_allDefaultBits_arr(m_arrdimSize)

    gL_UDRE_msbbit_arr = m_msbbit_arr
    gL_UDRE_lsbbit_arr = m_lsbbit_arr
    gL_UDRE_bitwidth_arr = m_bitwidth_arr
    gL_UDRE_DefaultOrReal_arr = m_DefaultOrReal_arr
    gL_UDRE_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_UDRE_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_UDRE_allDefaultBits_arr = m_allDefaultBits_arr
    gL_UDRE_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''201904 Ter
Public Function auto_eFuse_CMP_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_CMP_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    
    ''''201812XX
    ''''-------------------------------
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_tmpStrM As String
    Dim m_bitwidth As Long
    Dim m_lolmt_binarr() As Long
    Dim m_hilmt_binarr() As Long
    Dim m_allLoLMTBits_arr() As Long
    Dim m_allHiLMTBits_arr() As Long
    Dim m_catenameVbin As String
    Dim MaxLevelIndex As Long
    Dim m_Pmode As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    m_arrdimSize = gL_CMP_DigCapBits_Num - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(CMPFuse.Category)

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    
    ReDim m_allLoLMTBits_arr(m_arrdimSize)
    ReDim m_allHiLMTBits_arr(m_arrdimSize)
    
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)

    For i = 0 To m_size
        With CMPFuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth

            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_resolution = .Resoultion
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            ''''including "decimal", "default"
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then  '''' Job >= Stage
            If (m_msbbit_arr(i) < m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        End If
    Next i

    ReDim gL_CMP_msbbit_arr(m_size)
    ReDim gL_CMP_lsbbit_arr(m_size)
    ReDim gL_CMP_bitwidth_arr(m_size)
    ReDim gL_CMP_DefaultOrReal_arr(m_size)
    ReDim gL_CMP_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_CMP_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_CMP_allDefaultBits_arr(m_arrdimSize)

    gL_CMP_msbbit_arr = m_msbbit_arr
    gL_CMP_lsbbit_arr = m_lsbbit_arr
    gL_CMP_bitwidth_arr = m_bitwidth_arr
    gL_CMP_DefaultOrReal_arr = m_DefaultOrReal_arr
    gL_CMP_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_CMP_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_CMP_allDefaultBits_arr = m_allDefaultBits_arr
    gL_CMP_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''201904 Ter
Public Function auto_eFuse_CMPP_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_CMPP_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    
    ''''201812XX
    ''''-------------------------------
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_tmpStrM As String
    Dim m_bitwidth As Long
    Dim m_lolmt_binarr() As Long
    Dim m_hilmt_binarr() As Long
    Dim m_allLoLMTBits_arr() As Long
    Dim m_allHiLMTBits_arr() As Long
    Dim m_catenameVbin As String
    Dim MaxLevelIndex As Long
    Dim m_Pmode As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    m_arrdimSize = gL_CMPP_DigCapBits_Num - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(CMPP_Fuse.Category)

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    
    ReDim m_allLoLMTBits_arr(m_arrdimSize)
    ReDim m_allHiLMTBits_arr(m_arrdimSize)
    
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)

    For i = 0 To m_size
        With CMPP_Fuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth

            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_resolution = .Resoultion
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            ''''including "decimal", "default"
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then  '''' Job >= Stage
            If (m_msbbit_arr(i) < m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stageLEQjob_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        End If
    Next i

    ReDim gL_CMPP_msbbit_arr(m_size)
    ReDim gL_CMPP_lsbbit_arr(m_size)
    ReDim gL_CMPP_bitwidth_arr(m_size)
    ReDim gL_CMPP_DefaultOrReal_arr(m_size)
    ReDim gL_CMPP_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_CMPP_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_CMPP_allDefaultBits_arr(m_arrdimSize)

    gL_CMPP_msbbit_arr = m_msbbit_arr
    gL_CMPP_lsbbit_arr = m_lsbbit_arr
    gL_CMPP_bitwidth_arr = m_bitwidth_arr
    gL_CMPP_DefaultOrReal_arr = m_DefaultOrReal_arr
    gL_CMPP_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_CMPP_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_CMPP_allDefaultBits_arr = m_allDefaultBits_arr
    gL_CMPP_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''201904 Ter
Public Function auto_eFuse_CMPE_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_CMPE_glbVar_Init"

    Dim m_Site As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ''''-------------------------------
    Dim m_msbbit_arr() As Long
    Dim m_lsbbit_arr() As Long
    Dim m_bitwidth_arr() As Long
    Dim m_DefaultOrReal_arr() As Long
    Dim m_stage_bitFlag_arr() As Long
    Dim m_stage_early_bitFlag_arr() As Long
    Dim m_allDefaultBits_arr() As Long
    Dim m_DefValBitArr() As Long
    Dim m_size As Long
    Dim m_dfreal As String
    Dim m_stage As String
    Dim m_arrdimSize As Long
    
    ''''201812XX
    ''''-------------------------------
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_tmpStrM As String
    Dim m_bitwidth As Long
    Dim m_lolmt_binarr() As Long
    Dim m_hilmt_binarr() As Long
    Dim m_allLoLMTBits_arr() As Long
    Dim m_allHiLMTBits_arr() As Long
    Dim m_catenameVbin As String
    Dim MaxLevelIndex As Long
    Dim m_Pmode As Long
    Dim m_stageLEQjob_bitFlag_arr() As Long
    ''''-------------------------------
    
    m_arrdimSize = gL_CMPE_DigCapBits_Num - 1 '''' minus 1 is for the dimension of array
    m_size = UBound(CMPE_Fuse.Category)

    ReDim m_msbbit_arr(m_size)
    ReDim m_lsbbit_arr(m_size)
    ReDim m_bitwidth_arr(m_size)
    ReDim m_DefaultOrReal_arr(m_size)
    ReDim m_stage_bitFlag_arr(m_arrdimSize)
    ReDim m_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim m_allDefaultBits_arr(m_arrdimSize)
    
    ReDim m_allLoLMTBits_arr(m_arrdimSize)
    ReDim m_allHiLMTBits_arr(m_arrdimSize)
    
    ReDim m_stageLEQjob_bitFlag_arr(m_arrdimSize)

    For i = 0 To m_size
        With CMPE_Fuse.Category(i)
            m_msbbit_arr(i) = .MSBbit
            m_lsbbit_arr(i) = .LSBbit
            m_bitwidth_arr(i) = .BitWidth

            ReDim m_DefValBitArr(m_bitwidth_arr(i) - 1)
            m_DefValBitArr = .DefValBitArr
            m_dfreal = LCase(.Default_Real)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_resolution = .Resoultion
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
        End With

        If (m_dfreal = "real" Or m_dfreal = "bincut") Then
            m_DefaultOrReal_arr(i) = 0
        Else
            ''''including "decimal", "default"
            m_DefaultOrReal_arr(i) = 1 ''''default=1, real=0
        End If

        If (m_stage = gS_JobName + "_early") Then ''''Ex: cp1_early
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_early_bitFlag_arr(j) = 1
                Next j
            End If
        End If
        
        If (m_stage = gS_JobName) Then
            If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
                For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            Else
                For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                    m_stage_bitFlag_arr(j) = 1
                Next j
            End If
        End If

        ''''to get all DefaultBits PgmArr
        If (m_msbbit_arr(i) <= m_lsbbit_arr(i)) Then
            For j = m_msbbit_arr(i) To m_lsbbit_arr(i)
                k = m_lsbbit_arr(i) - j
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        Else
            For j = m_lsbbit_arr(i) To m_msbbit_arr(i)
                k = j - m_lsbbit_arr(i)
                m_allDefaultBits_arr(j) = m_DefValBitArr(k)
                ''''--------
                'm_allLoLMTBits_arr(j) = m_lolmt_binarr(k)
                'm_allHiLMTBits_arr(j) = m_hilmt_binarr(k)
            Next j
        End If
    Next i

    ReDim gL_CMPE_msbbit_arr(m_size)
    ReDim gL_CMPE_lsbbit_arr(m_size)
    ReDim gL_CMPE_bitwidth_arr(m_size)
    ReDim gL_CMPE_DefaultOrReal_arr(m_size)
    ReDim gL_CMPE_stage_bitFlag_arr(m_arrdimSize)
    ReDim gL_CMPE_stage_early_bitFlag_arr(m_arrdimSize)
    ReDim gL_CMPE_allDefaultBits_arr(m_arrdimSize)

    gL_CMPE_msbbit_arr = m_msbbit_arr
    gL_CMPE_lsbbit_arr = m_lsbbit_arr
    gL_CMPE_bitwidth_arr = m_bitwidth_arr
    gL_CMPE_DefaultOrReal_arr = m_DefaultOrReal_arr
    gL_CMPE_stage_bitFlag_arr = m_stage_bitFlag_arr
    gL_CMPE_stage_early_bitFlag_arr = m_stage_early_bitFlag_arr
    gL_CMPE_allDefaultBits_arr = m_allDefaultBits_arr
    gL_CMPE_stageLEQjob_bitFlag_arr = m_stageLEQjob_bitFlag_arr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20180711 New, 201811XX
''''This Function is just only to do once except for the re-validated / re-saved.
Public Function auto_eFuse_glbVar_Init() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_glbVar_Init"

    If (gB_findECID_flag = True) Then Call auto_eFuse_ECID_glbVar_Init

    If (gB_findCFG_flag = True) Then Call auto_eFuse_CFG_glbVar_Init
    
    ''201904 Ter
    If (gB_findUID_flag = True) Then Call auto_eFuse_UID_glbVar_Init
    
    If (gB_findUDR_flag = True) Then Call auto_eFuse_UDR_glbVar_Init
    
    If (gB_findUDRP_flag = True) Then Call auto_eFuse_UDRP_glbVar_Init
    
    If (gB_findUDRE_flag = True) Then Call auto_eFuse_UDRE_glbVar_Init
    
    If (gB_findSEN_flag = True) Then Call auto_eFuse_SEN_glbVar_Init
    
    If (gB_findMON_flag = True) Then Call auto_eFuse_MON_glbVar_Init
    
    If (gB_findCMP_flag = True) Then Call auto_eFuse_CMP_glbVar_Init
    
    If (gB_findCMPP_flag = True) Then Call auto_eFuse_CMPP_glbVar_Init
    
    If (gB_findCMPE_flag = True) Then Call auto_eFuse_CMPE_glbVar_Init

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20180711 New, 20181109
''''It will be called everytime in the function auto_eFuse_Initialize()
Public Function auto_eFuse_onProgramStarted() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_onProgramStarted"

    Dim m_Site As Variant

'Call auto_StartWatchTimer

    If (gL_1st_FuseSheetRead <> 0) Then Call auto_eFuse_Category_Reset

    ''''<MUST and Importance> pervent global DSP variable was reset
    'Debug.Print "gDL_eFuse_Orientation = " & gDL_eFuse_Orientation(0)
    'Debug.Print "gE_eFuse_Orientation = " & gE_eFuse_Orientation
    gDL_eFuse_Orientation = gE_eFuse_Orientation

'Call auto_StopWatchTimer("test__000")
    
    '''----------------------------------------
    '''Only Update for the Active Sites
    '''----------------------------------------

    '''Call auto_eFuse_Global_DSPWave_Reset ''''maybe could put outside, check it later
    
    ''''The below will copy dataArr to the DspWave (activeSites) only
    For Each m_Site In TheExec.sites
        ''''---------------------------------------------
        '''' Clear the global DSP DSPWave
        '''' it is also called in OnDSPGlobalVariableReset()
        ''''---------------------------------------------
        Call auto_eFuse_Global_DSPWave_Reset
        gDB_SerialType = False
'Call auto_StopWatchTimer("test__001")
        
        '''CRC calcBits
        gDW_ECID_CRC_calcBitsWave.Data = gL_ECID_CRC_calcBits
        gDW_CFG_CRC_calcBitsWave.Data = gL_CFG_CRC_calcBits

        '''ECID
        gDW_ECID_MSBBit_Cate.Data = gL_ECID_msbbit_arr
        gDW_ECID_LSBBit_Cate.Data = gL_ECID_lsbbit_arr
        gDW_ECID_BitWidth_Cate.Data = gL_ECID_bitwidth_arr
        gDW_ECID_DefaultReal_Cate.Data = gL_ECID_DefaultOrReal_arr
        gDW_ECID_Stage_BitFlag.Data = gL_ECID_stage_bitFlag_arr
        gDW_ECID_Stage_Early_BitFlag.Data = gL_ECID_stage_early_bitFlag_arr
        gDW_ECID_allDefaultBitWave.Data = gL_ECID_allDefaultBits_arr
        gDW_ECID_StageLEQJob_BitFlag.Data = gL_ECID_stageLEQjob_bitFlag_arr
        
        '''CFG
        gDW_CFG_MSBBit_Cate.Data = gL_CFG_msbbit_arr
        gDW_CFG_LSBBit_Cate.Data = gL_CFG_lsbbit_arr
        gDW_CFG_BitWidth_Cate.Data = gL_CFG_bitwidth_arr
        gDW_CFG_DefaultReal_Cate.Data = gL_CFG_DefaultOrReal_arr
        gDW_CFG_Stage_BitFlag.Data = gL_CFG_stage_bitFlag_arr
        gDW_CFG_Stage_Early_BitFlag.Data = gL_CFG_stage_early_bitFlag_arr
        gDW_CFG_Stage_Real_BitFlag.Data = gL_CFG_stage_real_bitFlag_arr
        gDW_CFG_allDefaultBitWave.Data = gL_CFG_allDefaultBits_arr
        gDW_CFG_StageLEQJob_BitFlag.Data = gL_CFG_stageLEQjob_bitFlag_arr
        gDW_CFG_SegFlag.Data = gL_CFG_SegFlag_arr
        
        ''201904 Ter
        ''UID
        gDW_UID_MSBBit_Cate.Data = gL_UID_msbbit_arr
        gDW_UID_LSBBit_Cate.Data = gL_UID_lsbbit_arr
        gDW_UID_BitWidth_Cate.Data = gL_UID_bitwidth_arr
        gDW_UID_DefaultReal_Cate.Data = gL_UID_DefaultOrReal_arr
        gDW_UID_Stage_BitFlag.Data = gL_UID_stage_bitFlag_arr
        gDW_UID_Stage_Early_BitFlag.Data = gL_UID_stage_early_bitFlag_arr
        gDW_UID_allDefaultBitWave.Data = gL_UID_allDefaultBits_arr
        gDW_UID_StageLEQJob_BitFlag.Data = gL_UID_stageLEQjob_bitFlag_arr
        
        ''UDR
        gDW_UDR_MSBBit_Cate.Data = gL_UDR_msbbit_arr
        gDW_UDR_LSBBit_Cate.Data = gL_UDR_lsbbit_arr
        gDW_UDR_BitWidth_Cate.Data = gL_UDR_bitwidth_arr
        gDW_UDR_DefaultReal_Cate.Data = gL_UDR_DefaultOrReal_arr
        gDW_UDR_Stage_BitFlag.Data = gL_UDR_stage_bitFlag_arr
        gDW_UDR_Stage_Early_BitFlag.Data = gL_UDR_stage_early_bitFlag_arr
        gDW_UDR_allDefaultBitWave.Data = gL_UDR_allDefaultBits_arr
        gDW_UDR_StageLEQJob_BitFlag.Data = gL_UDR_stageLEQjob_bitFlag_arr
        
        ''UDRP
        gDW_UDRP_MSBBit_Cate.Data = gL_UDRP_msbbit_arr
        gDW_UDRP_LSBBit_Cate.Data = gL_UDRP_lsbbit_arr
        gDW_UDRP_BitWidth_Cate.Data = gL_UDRP_bitwidth_arr
        gDW_UDRP_DefaultReal_Cate.Data = gL_UDRP_DefaultOrReal_arr
        gDW_UDRP_Stage_BitFlag.Data = gL_UDRP_stage_bitFlag_arr
        gDW_UDRP_Stage_Early_BitFlag.Data = gL_UDRP_stage_early_bitFlag_arr
        gDW_UDRP_allDefaultBitWave.Data = gL_UDRP_allDefaultBits_arr
        gDW_UDRP_StageLEQJob_BitFlag.Data = gL_UDRP_stageLEQjob_bitFlag_arr
        
        ''UDRE
        gDW_UDRE_MSBBit_Cate.Data = gL_UDRE_msbbit_arr
        gDW_UDRE_LSBBit_Cate.Data = gL_UDRE_lsbbit_arr
        gDW_UDRE_BitWidth_Cate.Data = gL_UDRE_bitwidth_arr
        gDW_UDRE_DefaultReal_Cate.Data = gL_UDRE_DefaultOrReal_arr
        gDW_UDRE_Stage_BitFlag.Data = gL_UDRE_stage_bitFlag_arr
        gDW_UDRE_Stage_Early_BitFlag.Data = gL_UDRE_stage_early_bitFlag_arr
        gDW_UDRE_allDefaultBitWave.Data = gL_UDRE_allDefaultBits_arr
        gDW_UDRE_StageLEQJob_BitFlag.Data = gL_UDRE_stageLEQjob_bitFlag_arr
        
        ''SEN
        gDW_SEN_MSBBit_Cate.Data = gL_SEN_msbbit_arr
        gDW_SEN_LSBBit_Cate.Data = gL_SEN_lsbbit_arr
        gDW_SEN_BitWidth_Cate.Data = gL_SEN_bitwidth_arr
        gDW_SEN_DefaultReal_Cate.Data = gL_SEN_DefaultOrReal_arr
        gDW_SEN_Stage_BitFlag.Data = gL_SEN_stage_bitFlag_arr
        gDW_SEN_Stage_Early_BitFlag.Data = gL_SEN_stage_early_bitFlag_arr
        gDW_SEN_allDefaultBitWave.Data = gL_SEN_allDefaultBits_arr
        gDW_SEN_StageLEQJob_BitFlag.Data = gL_SEN_stageLEQjob_bitFlag_arr
        
        ''MON
        gDW_MON_MSBBit_Cate.Data = gL_MON_msbbit_arr
        gDW_MON_LSBBit_Cate.Data = gL_MON_lsbbit_arr
        gDW_MON_BitWidth_Cate.Data = gL_MON_bitwidth_arr
        gDW_MON_DefaultReal_Cate.Data = gL_MON_DefaultOrReal_arr
        gDW_MON_Stage_BitFlag.Data = gL_MON_stage_bitFlag_arr
        gDW_MON_Stage_Early_BitFlag.Data = gL_MON_stage_early_bitFlag_arr
        gDW_MON_allDefaultBitWave.Data = gL_MON_allDefaultBits_arr
        gDW_MON_StageLEQJob_BitFlag.Data = gL_MON_stageLEQjob_bitFlag_arr
        
        ''CMP
        gDW_CMP_MSBBit_Cate.Data = gL_CMP_msbbit_arr
        gDW_CMP_LSBBit_Cate.Data = gL_CMP_lsbbit_arr
        gDW_CMP_BitWidth_Cate.Data = gL_CMP_bitwidth_arr
        gDW_CMP_DefaultReal_Cate.Data = gL_CMP_DefaultOrReal_arr
        gDW_CMP_Stage_BitFlag.Data = gL_CMP_stage_bitFlag_arr
        gDW_CMP_Stage_Early_BitFlag.Data = gL_CMP_stage_early_bitFlag_arr
        gDW_CMP_allDefaultBitWave.Data = gL_CMP_allDefaultBits_arr
        gDW_CMP_StageLEQJob_BitFlag.Data = gL_CMP_stageLEQjob_bitFlag_arr
        
        ''CMPP
        gDW_CMPP_MSBBit_Cate.Data = gL_CMPP_msbbit_arr
        gDW_CMPP_LSBBit_Cate.Data = gL_CMPP_lsbbit_arr
        gDW_CMPP_BitWidth_Cate.Data = gL_CMPP_bitwidth_arr
        gDW_CMPP_DefaultReal_Cate.Data = gL_CMPP_DefaultOrReal_arr
        gDW_CMPP_Stage_BitFlag.Data = gL_CMPP_stage_bitFlag_arr
        gDW_CMPP_Stage_Early_BitFlag.Data = gL_CMPP_stage_early_bitFlag_arr
        gDW_CMPP_allDefaultBitWave.Data = gL_CMPP_allDefaultBits_arr
        gDW_CMPP_StageLEQJob_BitFlag.Data = gL_CMPP_stageLEQjob_bitFlag_arr
        
        ''CMPE
        gDW_CMPE_MSBBit_Cate.Data = gL_CMPE_msbbit_arr
        gDW_CMPE_LSBBit_Cate.Data = gL_CMPE_lsbbit_arr
        gDW_CMPE_BitWidth_Cate.Data = gL_CMPE_bitwidth_arr
        gDW_CMPE_DefaultReal_Cate.Data = gL_CMPE_DefaultOrReal_arr
        gDW_CMPE_Stage_BitFlag.Data = gL_CMPE_stage_bitFlag_arr
        gDW_CMPE_Stage_Early_BitFlag.Data = gL_CMPE_stage_early_bitFlag_arr
        gDW_CMPE_allDefaultBitWave.Data = gL_CMPE_allDefaultBits_arr
        gDW_CMPE_StageLEQJob_BitFlag.Data = gL_CMPE_stageLEQjob_bitFlag_arr

        gDW_MON_CRC_calcBitsWave = gDW_MON_CRC_calcBits_Temp.Copy
        gDW_SEN_CRC_calcBitsWave = gDW_SEN_CRC_calcBits_Temp.Copy
'Call auto_StopWatchTimer("test__002")

    Next m_Site
'Call auto_StopWatchTimer("test__002")

'    gDW_UID_CRC_calcBitsWave = gDW_UID_CRC_calcBits_Temp.Copy
'    gDW_MON_CRC_calcBitsWave = gDW_MON_CRC_calcBits_Temp.Copy
'    gDW_SEN_CRC_calcBitsWave = gDW_SEN_CRC_calcBits_Temp.Copy
  
    ''''initialize the below variable for each run (initFlows)
    gB_ECID_decode_flag = False
    gB_CFG_decode_flag = False
    gB_UDR_decode_flag = False
    gB_SEN_decode_flag = False
    gB_MON_decode_flag = False
    gB_CMP_decode_flag = False

    HramLotId = ""
    HramWaferId = 0
    HramXCoord = -32768
    HramYCoord = -32768
    gS_ECID_CRC_HexStr = "" ''''MUST be, 20161004 update
    gS_CFG_CRC_HexStr = "" ''''MUST be
    gS_MON_CRC_HexStr = "" ''''MUST be
    gS_SEN_CRC_HexStr = "" ''''MUST be

    'Set eFuse Global Data initial
    gB_CFGSVM_BIT_Read_ValueisONE = False ''''<MUST>

    gS_USI_BitStr = ""
    gS_UDRE_USI_BitStr = ""
    gS_UDRP_USI_BitStr = ""

    ''''<MUST> Reset again, because it was changed in the function auto_Chk_ECID_Content_DEID().
    ''''<Notice and Be careful>
    XCOORD_LoLMT = ECIDFuse.Category(ECIDIndex("X_Coordinate")).LoLMT
    XCOORD_HiLMT = ECIDFuse.Category(ECIDIndex("X_Coordinate")).HiLMT
    YCOORD_LoLMT = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).LoLMT
    YCOORD_HiLMT = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).HiLMT
'Call auto_StopWatchTimer("test__003")

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20180711 New, only reset all Read categories and Write categories of "Real and BinCut"
Public Function auto_eFuse_Category_Reset() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Category_Reset"
    
    Dim m_Site As Variant
    Dim i, k As Long
    Dim m_defreal As String
    Dim m_bitwidth As Long
    Dim m_defval As Variant
    Dim ms_defval As New SiteVariant
    Dim m_binarr() As Long
    Dim m_catename As String
    Dim m_algorithm As String
    ''Dim m_zeroWave As New DSPWave

    Dim m_resetFuseParam As EFuseCategoryParamResultSyntax
    
    ''''By this way to let all Fuse parameters to Nothing/CLear
    ''''Could choose any one of the members as the representative (.Decimal, .HexStr, ...)
    Set m_resetFuseParam.Decimal = Nothing

    If (gB_findECID_flag = True) Then
        For i = 0 To UBound(ECIDFuse.Category)
            m_bitwidth = ECIDFuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(ECIDFuse.Category(i).Default_Real)
            ECIDFuse.Category(i).PatTestPass_Flag = True
            ECIDFuse.Category(i).Read = m_resetFuseParam
            m_algorithm = ECIDFuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                m_defval = auto_checkDefaultValue(0, m_algorithm, m_binarr, m_bitwidth, m_defreal)
                m_catename = ECIDFuse.Category(i).Name
                ECIDFuse.Category(i).Write = m_resetFuseParam
                ms_defval = m_defval
                Call auto_eFuse_SetWriteVariable_SiteAware("ECID", m_catename, ms_defval, False)
            End If
        Next i
    End If

    If (gB_findCFG_flag = True) Then
        For i = 0 To UBound(CFGFuse.Category)
            m_bitwidth = CFGFuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(CFGFuse.Category(i).Default_Real)
            CFGFuse.Category(i).PatTestPass_Flag = True
            CFGFuse.Category(i).Read = m_resetFuseParam
            m_algorithm = CFGFuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                m_catename = CFGFuse.Category(i).Name
                CFGFuse.Category(i).Write = m_resetFuseParam
                m_defval = auto_checkDefaultValue(0, m_algorithm, m_binarr, m_bitwidth, m_defreal)
                ms_defval = m_defval
                Call auto_eFuse_SetWriteVariable_SiteAware("CFG", m_catename, ms_defval, False)

            End If
        Next i
    End If

    If (gB_findUID_flag = True) Then
        For i = 0 To UBound(UIDFuse.Category)
            m_bitwidth = UIDFuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(UIDFuse.Category(i).Default_Real)
            UIDFuse.Category(i).PatTestPass_Flag = True
            UIDFuse.Category(i).Read = m_resetFuseParam
            m_algorithm = UIDFuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                m_catename = UIDFuse.Category(i).Name
                UIDFuse.Category(i).Write = m_resetFuseParam
                m_defval = auto_checkDefaultValue(0, m_algorithm, m_binarr, m_bitwidth, m_defreal)
                ms_defval = m_defval
                Call auto_eFuse_SetWriteVariable_SiteAware("UID", m_catename, ms_defval, False)
            End If
        Next i
    End If

    If (gB_findUDR_flag = True) Then
        For i = 0 To UBound(UDRFuse.Category)
            m_bitwidth = UDRFuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(UDRFuse.Category(i).Default_Real)
            UDRFuse.Category(i).PatTestPass_Flag = True
            UDRFuse.Category(i).Read = m_resetFuseParam
            m_algorithm = UDRFuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                m_catename = UDRFuse.Category(i).Name
                UDRFuse.Category(i).Write = m_resetFuseParam
                ''20191221 modify central update
                m_defval = auto_checkDefaultValue(0, m_algorithm, m_binarr, m_bitwidth, m_defreal)           ''' added on 2019/10/28 for fix input values not initialized for DSPSrc
                'UDRFuse.Category(i).DefValBitArr = m_binarr                                                    ''' added on 2019/10/28 for fix input values not initialized for DSPSrc
                ms_defval = m_defval                                                                                ''' added on 2019/10/28 for fix input values not initialized for DSPSrc
                Call auto_eFuse_SetWriteVariable_SiteAware("UDR", m_catename, ms_defval, False)     ' added on 2019/10/28 for fix input values not initialized for DSPSrc
            End If
        Next i
    End If

    If (gB_findCMP_flag = True) Then
        For i = 0 To UBound(CMPFuse.Category)
            m_bitwidth = CMPFuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(CMPFuse.Category(i).Default_Real)
            CMPFuse.Category(i).PatTestPass_Flag = True
            CMPFuse.Category(i).Read = m_resetFuseParam
            m_algorithm = CMPFuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                CMPFuse.Category(i).Write = m_resetFuseParam
            End If
        Next i
    End If

    If (gB_findSEN_flag = True) Then
        For i = 0 To UBound(SENFuse.Category)
            m_bitwidth = SENFuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(SENFuse.Category(i).Default_Real)
            SENFuse.Category(i).PatTestPass_Flag = True
            SENFuse.Category(i).Read = m_resetFuseParam
            m_algorithm = SENFuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                SENFuse.Category(i).Write = m_resetFuseParam
            End If
        Next i
    End If
    
    If (gB_findMON_flag = True) Then
        For i = 0 To UBound(MONFuse.Category)
            m_bitwidth = MONFuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(MONFuse.Category(i).Default_Real)
            MONFuse.Category(i).PatTestPass_Flag = True
            MONFuse.Category(i).Read = m_resetFuseParam
            m_algorithm = MONFuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                m_catename = MONFuse.Category(i).Name
                MONFuse.Category(i).Write = m_resetFuseParam
                m_defval = auto_checkDefaultValue(0, m_algorithm, m_binarr, m_bitwidth, m_defreal)
                ms_defval = m_defval
                Call auto_eFuse_SetWriteVariable_SiteAware("MON", m_catename, ms_defval, False)

            End If
        Next i
    End If

    ''''20171103 update
    If (gB_findUDRE_flag = True) Then
        For i = 0 To UBound(UDRE_Fuse.Category)
            m_bitwidth = UDRE_Fuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(UDRE_Fuse.Category(i).Default_Real)
            UDRE_Fuse.Category(i).PatTestPass_Flag = True
            UDRE_Fuse.Category(i).Read = m_resetFuseParam
            m_algorithm = UDRE_Fuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                m_catename = UDRE_Fuse.Category(i).Name
                UDRE_Fuse.Category(i).Write = m_resetFuseParam
                m_defval = auto_checkDefaultValue(0, m_algorithm, m_binarr, m_bitwidth, m_defreal)
                ms_defval = m_defval
                Call auto_eFuse_SetWriteVariable_SiteAware("UDRE", m_catename, ms_defval, False)

            End If
        Next i
    End If

    If (gB_findUDRP_flag = True) Then
        For i = 0 To UBound(UDRP_Fuse.Category)
            m_bitwidth = UDRP_Fuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
            UDRP_Fuse.Category(i).PatTestPass_Flag = True
            UDRP_Fuse.Category(i).Read = m_resetFuseParam
            m_algorithm = UDRP_Fuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                m_catename = UDRP_Fuse.Category(i).Name
                UDRP_Fuse.Category(i).Write = m_resetFuseParam
                m_defval = auto_checkDefaultValue(0, m_algorithm, m_binarr, m_bitwidth, m_defreal)
                ms_defval = m_defval
                Call auto_eFuse_SetWriteVariable_SiteAware("UDRP", m_catename, ms_defval, False)

            End If
        Next i
    End If
    
    If (gB_findCMPE_flag = True) Then
        For i = 0 To UBound(CMPE_Fuse.Category)
            m_bitwidth = CMPE_Fuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(CMPE_Fuse.Category(i).Default_Real)
            CMPE_Fuse.Category(i).PatTestPass_Flag = True
            CMPE_Fuse.Category(i).Read = m_resetFuseParam
            m_algorithm = CMPE_Fuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                CMPE_Fuse.Category(i).Write = m_resetFuseParam
            End If
        Next i
    End If

    If (gB_findCMPP_flag = True) Then
        For i = 0 To UBound(CMPP_Fuse.Category)
            m_bitwidth = CMPP_Fuse.Category(i).BitWidth
            ''m_zeroWave.CreateConstant 0, m_bitwidth, DspLong
            m_defreal = LCase(CMPP_Fuse.Category(i).Default_Real)
            CMPP_Fuse.Category(i).PatTestPass_Flag = True
            CMPP_Fuse.Category(i).Read = m_resetFuseParam
            m_algorithm = CMPP_Fuse.Category(i).algorithm
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                CMPP_Fuse.Category(i).Write = m_resetFuseParam
            End If
        Next i
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150625 New Function
Public Function auto_eFuse_SetPatTestPass_Flag_SiteAware(ByVal FuseType As String, m_catename As String, m_flag As SiteBoolean, Optional showPrint As Boolean = True) As SiteVariant
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_SetPatTestPass_Flag_SiteAware"

    Dim m_len As Long
    Dim m_dlogstr As String
    Dim m_Site As Variant

    FuseType = UCase(Trim(FuseType))
    
    If (FuseType = "ECID") Then
        ECIDFuse.Category(ECIDIndex(m_catename)).PatTestPass_Flag = m_flag

    ElseIf (FuseType = "CFG") Then
        CFGFuse.Category(CFGIndex(m_catename)).PatTestPass_Flag = m_flag

    ElseIf (FuseType = "UID") Then
        UIDFuse.Category(UIDIndex(m_catename)).PatTestPass_Flag = m_flag

    ElseIf (FuseType = "UDR") Then
        UDRFuse.Category(UDRIndex(m_catename)).PatTestPass_Flag = m_flag

    ElseIf (FuseType = "SEN") Then
        SENFuse.Category(SENIndex(m_catename)).PatTestPass_Flag = m_flag
        
    ElseIf (FuseType = "MON") Then
        MONFuse.Category(MONIndex(m_catename)).PatTestPass_Flag = m_flag

    ElseIf (FuseType = "CMP") Then
        CMPFuse.Category(CMPIndex(m_catename)).PatTestPass_Flag = m_flag

    ElseIf (FuseType = "UDRE") Then
        UDRE_Fuse.Category(UDRE_Index(m_catename)).PatTestPass_Flag = m_flag

    ElseIf (FuseType = "UDRP") Then
        UDRP_Fuse.Category(UDRP_Index(m_catename)).PatTestPass_Flag = m_flag

    ElseIf (FuseType = "CMPE") Then
        CMPE_Fuse.Category(CMPE_Index(m_catename)).PatTestPass_Flag = m_flag

    ElseIf (FuseType = "CMPP") Then
        CMPP_Fuse.Category(CMPP_Index(m_catename)).PatTestPass_Flag = m_flag

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If
    
    Set auto_eFuse_SetPatTestPass_Flag_SiteAware = m_flag
    
    If (showPrint) Then
        Dim fusetype_org As String
       
        For Each m_Site In TheExec.sites
            fusetype_org = FormatNumeric(FuseType, 4)
            fusetype_org = "Site(" + CStr(m_Site) + ") " + fusetype_org + FormatNumeric("Fuse SetPatTestPass_Flag", -35)
            m_dlogstr = vbTab & FormatNumeric(fusetype_org, Len(fusetype_org)) + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_flag, -10)
            TheExec.Datalog.WriteComment m_dlogstr
            fusetype_org = ""
            m_dlogstr = ""
        Next m_Site
    End If
    
'    If (showPrint) Then
'        Dim fusetype_org As String
'        'm_len = auto_eFuse_GetCatenameMaxLen(FuseType)
'        fusetype_org = FormatNumeric(FuseType, 4)
'        For Each m_Site In TheExec.Sites
'            'FuseType = "Site(" + CStr(m_Site) + ") " + fusetype_org + FormatNumeric("Fuse SetPatTestPass_Flag", -25)
'            FuseType = FormatNumeric(FuseType, 4)
'            FuseType = "Site(" + CStr(m_Site) + ") " + FuseType + FormatNumeric("Fuse SetPatTestPass_Flag", -25)
'
'            m_dlogstr = vbTab & FormatNumeric(FuseType, Len(FuseType)) + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_flag, -5)
'            TheExec.Datalog.WriteComment m_dlogstr
'            FuseType = ""
'            m_dlogstr = ""
'        Next m_Site
'    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20180726 Update Function to SiteAware
Public Function auto_eFuse_SetWriteVariable_SiteAware(ByVal FuseType As String, m_catename As String, ByVal in_value As SiteVariant, _
                                           Optional showPrint As Boolean = False) As SiteVariant

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_SetWriteVariable_SiteAware"

    Dim m_len As Long
    Dim m_decimal As Variant
    Dim m_dlogstr As String
    Dim m_Site As Variant

    Dim m_BinStr As String
    Dim m_HexStr As String
    Dim m_idx As Long
    Dim m_bitStrM As String
    Dim m_binarr() As Long
    Dim m_bitwidth As Long
    Dim m_value As Variant
    Dim m_tmpWave As New DSPWave
    Dim is_sameValue_Flag As Boolean
    Dim fusetype_org As String
    
    ''''Check if the Value of all Sites are same.
    is_sameValue_Flag = auto_is_sameValue_Sites(in_value)

    Dim m_FuseWrite As EFuseCategoryParamResultSyntax

    FuseType = UCase(Trim(FuseType))
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)

    If (FuseType = "ECID") Then
        m_idx = ECIDIndex(m_catename)
        m_bitwidth = ECIDFuse.Category(m_idx).BitWidth
        m_FuseWrite = ECIDFuse.Category(m_idx).Write
    ElseIf (FuseType = "CFG") Then
        m_idx = CFGIndex(m_catename)
        m_bitwidth = CFGFuse.Category(m_idx).BitWidth
        m_FuseWrite = CFGFuse.Category(m_idx).Write
    ElseIf (FuseType = "UID") Then
        m_idx = UIDIndex(m_catename)
        m_bitwidth = UIDFuse.Category(m_idx).BitWidth
        m_FuseWrite = UIDFuse.Category(m_idx).Write
    ElseIf (FuseType = "UDR") Then
        m_idx = UDRIndex(m_catename)
        m_bitwidth = UDRFuse.Category(m_idx).BitWidth
        m_FuseWrite = UDRFuse.Category(m_idx).Write
    ElseIf (FuseType = "SEN") Then
        m_idx = SENIndex(m_catename)
        m_bitwidth = SENFuse.Category(m_idx).BitWidth
        m_FuseWrite = SENFuse.Category(m_idx).Write
    ElseIf (FuseType = "MON") Then
        m_idx = MONIndex(m_catename)
        m_bitwidth = MONFuse.Category(m_idx).BitWidth
        m_FuseWrite = MONFuse.Category(m_idx).Write
    ElseIf (FuseType = "CMP") Then
        m_idx = CMPIndex(m_catename)
        m_bitwidth = CMPFuse.Category(m_idx).BitWidth
        m_FuseWrite = CMPFuse.Category(m_idx).Write
    ElseIf (FuseType = "UDRE") Then
        m_idx = UDRE_Index(m_catename)
        m_bitwidth = UDRE_Fuse.Category(m_idx).BitWidth
        m_FuseWrite = UDRE_Fuse.Category(m_idx).Write
    ElseIf (FuseType = "UDRP") Then
        m_idx = UDRP_Index(m_catename)
        m_bitwidth = UDRP_Fuse.Category(m_idx).BitWidth
        m_FuseWrite = UDRP_Fuse.Category(m_idx).Write
    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP)"
        GoTo errHandler
        ''''nothing
    End If

    If (is_sameValue_Flag = True) Then
        ''''Only do once for the same value in all sites
        For Each m_Site In TheExec.sites
            m_value = in_value(m_Site)
            Exit For
        Next m_Site
        
        ''''20160620 update, if it's Hex, it MUST be with the prefix "0x" or "x"
        If (auto_isHexString(CStr(m_value))) Then
            'm_hexStr = m_value
            If (auto_chkHexStr_isOver7FFFFFFF(CStr(m_value)) = False) Then
                m_value = Replace(UCase(CStr(m_value)), "0X", "", 1, 1)
                m_value = Replace(UCase(CStr(m_value)), "X", "", 1, 1) ''''20171211 update
                m_value = Replace(UCase(CStr(m_value)), "_", "")       ''''20171211 update
                m_HexStr = "0x" + CStr(m_value)
                m_value = CLng("&H" & m_value) ''''Here it's Hex2Dec
                'm_decimal = m_value
            Else
                ''''<MUST> keep prefix "0x" or "x"
                m_value = Replace(UCase(CStr(m_value)), "0X", "", 1, 1)
                m_value = Replace(UCase(CStr(m_value)), "X", "", 1, 1) ''''20171211 update
                m_value = Replace(UCase(CStr(m_value)), "_", "")       ''''20171211 update
                m_HexStr = "0x" + CStr(m_value)
            End If
            m_decimal = m_value
    
        ElseIf (auto_isBinaryString(CStr(m_value))) Then ''''20171211 add
            m_BinStr = Replace(UCase(CStr(m_value)), "B", "", 1, 1) ''''remove the first "B" character
            m_BinStr = Replace(m_BinStr, "_", "")                   ''''for case like as b0001_0101
            
            ''''convert to Hex with the prefix "0x"
            m_HexStr = "0x" + auto_BinStr2HexStr(m_BinStr, 1)
            If (auto_chkHexStr_isOver7FFFFFFF(CStr(m_HexStr)) = False) Then
                m_HexStr = Replace(UCase(CStr(m_HexStr)), "0X", "", 1, 1)
                m_value = CLng("&H" & m_HexStr) ''''Here it's Hex2Dec
                m_decimal = m_value
                m_HexStr = "0x" + m_HexStr
            Else
                ''''<MUST> keep prefix "0x" or "x"
                m_decimal = m_HexStr
            End If
            'm_hexStr = "0x" + m_hexStr
        Else
            m_decimal = m_value 'CLng(m_value), 20170911 update for the value over 31 bits
            m_HexStr = auto_Value2HexStr(m_decimal, m_bitwidth)
        End If
            
        ''''20171211 update
        If (m_decimal < 0) Then
            GoTo errHandler
        End If
        
        If m_decimal > (2 ^ m_bitwidth - 1) And m_bitwidth < 32 Then
            m_decimal = 0
            m_HexStr = "0x0"
            funcName = funcName + ": HIP/IDS/vdd-binning has been overflow "
            GoTo errHandler
        End If
            
        ''''20180711 New for DSPWave
        m_bitStrM = auto_HexStr2BinStr_EFUSE(m_HexStr, m_bitwidth, m_binarr)
        m_tmpWave.CreateConstant 0, m_bitwidth, DspLong
        m_tmpWave.Data = m_binarr

        With m_FuseWrite
            .Decimal = m_decimal
            .BitArrWave = m_tmpWave.Copy
            .HexStr = m_HexStr
            .Value = m_decimal
            .BitSummation = m_tmpWave.CalcSum
            .BitStrM = m_bitStrM
            .BitStrL = StrReverse(m_bitStrM)
            .ValStr = CStr(m_decimal)
        End With

        If (FuseType = "ECID") Then
            ECIDFuse.Category(m_idx).Write = m_FuseWrite

        ElseIf (FuseType = "CFG") Then
            CFGFuse.Category(m_idx).Write = m_FuseWrite

        ElseIf (FuseType = "UID") Then
            UIDFuse.Category(m_idx).Write = m_FuseWrite

        ElseIf (FuseType = "UDR") Then
            UDRFuse.Category(m_idx).Write = m_FuseWrite

        ElseIf (FuseType = "SEN") Then
            SENFuse.Category(m_idx).Write = m_FuseWrite

        ElseIf (FuseType = "MON") Then
            MONFuse.Category(m_idx).Write = m_FuseWrite

        ElseIf (FuseType = "CMP") Then
            CMPFuse.Category(m_idx).Write = m_FuseWrite

        ElseIf (FuseType = "UDRE") Then
            UDRE_Fuse.Category(m_idx).Write = m_FuseWrite

        ElseIf (FuseType = "UDRP") Then
            UDRP_Fuse.Category(m_idx).Write = m_FuseWrite

        Else
            TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP)"
            GoTo errHandler
            ''''nothing
        End If
        
        If (showPrint) Then
            'Dim fusetype_org As String
            For Each m_Site In TheExec.sites
            fusetype_org = FormatNumeric(FuseType, 4)
            m_dlogstr = vbTab & "Site(" + CStr(m_Site) + ") " + fusetype_org + FormatNumeric("Fuse SetWriteVariable_SiteAware", -35)
            m_dlogstr = m_dlogstr + FormatNumeric(m_catename, Len(fusetype_org)) + " = " + FormatNumeric(m_decimal, -10)
            TheExec.Datalog.WriteComment m_dlogstr
            m_dlogstr = ""
            fusetype_org = ""
            Next m_Site
        End If
        
''        ''''20171211 update
''        If (m_decimal < 0) Then
''            GoTo errHandler
''        End If

    Else
        ''''In the Site Iteration
        For Each m_Site In TheExec.sites
            m_value = in_value(m_Site)

            ''''20160620 update, if it's Hex, it MUST be with the prefix "0x" or "x"
            If (auto_isHexString(CStr(m_value))) Then
                'm_hexStr = m_value
                If (auto_chkHexStr_isOver7FFFFFFF(CStr(m_value)) = False) Then
                    m_value = Replace(UCase(CStr(m_value)), "0X", "", 1, 1)
                    m_value = Replace(UCase(CStr(m_value)), "X", "", 1, 1) ''''20171211 update
                    m_value = Replace(UCase(CStr(m_value)), "_", "")       ''''20171211 update
                    m_HexStr = "0x" + CStr(m_value)
                    m_value = CLng("&H" & m_value) ''''Here it's Hex2Dec
                Else
                    ''''<MUST> keep prefix "0x" or "x"
                    m_value = Replace(UCase(CStr(m_value)), "0X", "", 1, 1)
                    m_value = Replace(UCase(CStr(m_value)), "X", "", 1, 1) ''''20171211 update
                    m_value = Replace(UCase(CStr(m_value)), "_", "")       ''''20171211 update
                    m_HexStr = "0x" + CStr(m_value)
                End If
                m_decimal = m_value
        
            ElseIf (auto_isBinaryString(CStr(m_value))) Then ''''20171211 add
                m_BinStr = Replace(UCase(CStr(m_value)), "B", "", 1, 1) ''''remove the first "B" character
                m_BinStr = Replace(m_BinStr, "_", "")                   ''''for case like as b0001_0101
                
                ''''convert to Hex with the prefix "0x"
                m_HexStr = "0x" + auto_BinStr2HexStr(m_BinStr, 1)
                If (auto_chkHexStr_isOver7FFFFFFF(CStr(m_HexStr)) = False) Then
                    m_HexStr = Replace(UCase(CStr(m_HexStr)), "0X", "", 1, 1)
                    m_value = CLng("&H" & m_HexStr) ''''Here it's Hex2Dec
                    m_decimal = m_value
                    m_HexStr = "0x" + m_HexStr
                Else
                    ''''<MUST> keep prefix "0x" or "x"
                    m_decimal = m_HexStr
                End If
                'm_hexStr = "0x" + m_hexStr
            Else
                m_decimal = m_value 'CLng(m_value), 20170911 update for the value over 31 bits
                m_HexStr = auto_Value2HexStr(m_decimal, m_bitwidth)
            End If

            ''''20171211 update
            If (m_decimal < 0) Then
                GoTo errHandler
            End If
            
            If m_decimal > (2 ^ m_bitwidth - 1) Then
                m_decimal = 0
                m_HexStr = "0x0"
                funcName = funcName + ": HIP/IDS/vdd-binning has been overflow "
                GoTo errHandler
            End If

            ''''20180711 New for DSPWave
            m_bitStrM = auto_HexStr2BinStr_EFUSE(m_HexStr, m_bitwidth, m_binarr)
            m_tmpWave.CreateConstant 0, m_bitwidth, DspLong
            m_tmpWave(m_Site).Data = m_binarr
    
            With m_FuseWrite
                .Decimal(m_Site) = m_decimal
                .BitArrWave = m_tmpWave.Copy
                .HexStr(m_Site) = m_HexStr
                .Value(m_Site) = m_decimal
                .BitSummation(m_Site) = m_tmpWave.CalcSum
                .BitStrM(m_Site) = m_bitStrM
                .BitStrL(m_Site) = StrReverse(m_bitStrM)
                .ValStr(m_Site) = CStr(m_decimal)
            End With
    
            If (FuseType = "ECID") Then
                ECIDFuse.Category(m_idx).Write = m_FuseWrite
    
            ElseIf (FuseType = "CFG") Then
                CFGFuse.Category(m_idx).Write = m_FuseWrite
    
            ElseIf (FuseType = "UID") Then
                UIDFuse.Category(m_idx).Write = m_FuseWrite
    
            ElseIf (FuseType = "UDR") Then
                UDRFuse.Category(m_idx).Write = m_FuseWrite
    
            ElseIf (FuseType = "SEN") Then
                SENFuse.Category(m_idx).Write = m_FuseWrite
    
            ElseIf (FuseType = "MON") Then
                MONFuse.Category(m_idx).Write = m_FuseWrite
    
            ElseIf (FuseType = "CMP") Then
                CMPFuse.Category(m_idx).Write = m_FuseWrite
    
            ElseIf (FuseType = "UDRE") Then
                UDRE_Fuse.Category(m_idx).Write = m_FuseWrite
    
            ElseIf (FuseType = "UDRP") Then
                UDRP_Fuse.Category(m_idx).Write = m_FuseWrite
    
            Else
                TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,SEN,MON,CMP,UDRE,UDRP)"
                GoTo errHandler
                ''''nothing
            End If
            
            If (showPrint) Then
                'Dim m_FuseStr As String
                fusetype_org = FormatNumeric(FuseType, 4)
                m_dlogstr = vbTab & "Site(" + CStr(m_Site) + ") " + fusetype_org + FormatNumeric("Fuse SetWriteVariable_SiteAware", -35)
                m_dlogstr = m_dlogstr + FormatNumeric(m_catename, Len(fusetype_org)) + " = " + FormatNumeric(m_decimal, -10)
                TheExec.Datalog.WriteComment m_dlogstr
                 m_dlogstr = ""
            End If
            
'            If (showPrint) Then
'                FuseType = FormatNumeric(FuseType, 4)
'                m_dlogstr = vbTab & "Site(" + CStr(m_Site) + ") " + FuseType + FormatNumeric("Fuse SetWriteVariable_SiteAware", -35)
'                m_dlogstr = m_dlogstr + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_decimal, -10)
'                TheExec.Datalog.WriteComment m_dlogstr
'            End If
            
''            ''''20171211 update
''            If (m_decimal < 0) Then
''                GoTo errHandler
''            End If
        Next m_Site
    End If

    Set auto_eFuse_SetWriteVariable_SiteAware = m_FuseWrite.Decimal

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


''''20180726 Update Function to SiteAware
Public Function auto_is_sameValue_Sites(ByVal in_value As SiteVariant, Optional showPrint As Boolean = False) As Boolean

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_is_sameValue_Sites"

    Dim m_Site As Variant
    Dim m_tmpValue_ref As Variant
    Dim m_tmpValue As Variant
    
    auto_is_sameValue_Sites = True

    ''''Check if the Value of all Sites are same.
    For Each m_Site In TheExec.sites
        m_tmpValue_ref = in_value(m_Site)
        Exit For
    Next m_Site
    
    ''The below can not judge string case
    ''auto_is_sameValue_Sites = in_value.compare(EqualTo, m_tmpValue_ref).All(True)
    
    For Each m_Site In TheExec.sites
        If (in_value(m_Site) <> m_tmpValue_ref) Then
            auto_is_sameValue_Sites = False
            Exit For
        End If
    Next m_Site
    
    For Each m_Site In TheExec.sites
        If (in_value(m_Site) = "") Then in_value(m_Site) = 0
    Next m_Site
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''201808XX New, Here inWave is "doubleBitsWave" (effective pgm bits)
Public Function auto_eFuse_print_PgmBitsWave_Category(ByVal FuseType As eFuseBlockType, ByVal InWave As DSPWave, Optional showPrint As Boolean = True) As Boolean

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_print_PgmBitsWave_Category"

    If (showPrint = False) Then Exit Function
    
    Dim i As Long, j As Long, k As Long, bcnt As Long
    Dim m_Site As Variant
    Dim m_pgmBitArr() As Long
    Dim m_tmpStr As String
    Dim m_stage As String
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_MSBBit As Long
    Dim m_LSBbit As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim tmpdlgStr As String
    Dim m_bitstrL As String
    Dim m_bitStrM As String
    Dim m_tmpVal As Variant
    Dim m_HexStr As String
    Dim m_pgmlotidStr As String
    Dim m_matchFlag As Boolean
    Dim m_tmpbitStrM As String
    Dim m_resolution As Double
    Dim m_value As Variant
    Dim m_tmpStr1 As String
    Dim m_Fuseblock As EFuseCategorySyntax
    Dim m_FieldStr As String

    If (FuseType = eFuse_ECID) Then
        For Each m_Site In TheExec.sites
            m_pgmBitArr = InWave(m_Site).Data
            TheExec.Datalog.WriteComment ""
            m_FieldStr = ""
            For i = 0 To UBound(ECIDFuse.Category) - 1
                With ECIDFuse.Category(i)
                    m_catename = .Name
                    m_stage = LCase(.Stage)
                    m_algorithm = LCase(.algorithm)
                    m_MSBBit = .MSBbit
                    m_LSBbit = .LSBbit
                    m_bitwidth = .BitWidth
                    m_defval = .DefaultValue
                    m_defreal = LCase(.Default_Real)
                End With
    
                If (m_stage = gS_JobName) Then
                    ''''PgmBit datalog format
                    tmpdlgStr = "Site(" + CStr(m_Site) + ") Programming : " + FormatNumeric(m_catename, gI_ECID_catename_maxLen)
                    If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
                        If (m_algorithm = "crc") Then
                        Else
                            tmpdlgStr = tmpdlgStr + " [(LSB)" + Format(m_LSBbit, "0000") + ":" + Format(m_MSBBit, "0000") + "(MSB)] = "
                        End If
                    Else
                        tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "(LSB)] = "
                    End If
                    
                    m_bitStrM = ""
                    For j = 0 To m_bitwidth - 1
                        bcnt = m_MSBBit + j
                        m_bitStrM = m_bitStrM + CStr(m_pgmBitArr(bcnt))
                    Next j
                    m_bitstrL = StrReverse(m_bitStrM)
'                    If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
'                        m_FieldStr = m_bitstrL
'                    Else
'                        m_FieldStr = m_bitStrM
'                    End If
'                    m_hexStr = auto_Value2HexStr("b" + m_FieldStr, m_bitwidth)
                    m_HexStr = auto_Value2HexStr("b" + m_bitStrM, m_bitwidth)
                    
                    
                    If (m_algorithm = "lotid") Then
                        m_pgmlotidStr = ""
                        For j = 0 To EcidCharPerLotId - 1
                            m_tmpStr = ""
                            m_tmpStr = Mid(m_bitStrM, 1 + j * EcidBitPerLotIdChar, EcidBitPerLotIdChar) ''''EcidBitPerLotIdChar=6
                            m_pgmlotidStr = m_pgmlotidStr + auto_MappingBinStrtoChar(m_tmpStr)
                        Next j
                        m_decimal = m_pgmlotidStr
                    
                    ElseIf (m_algorithm = "crc") Then
                        If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
                            tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_LSBbit, "0000") + ":" + Format(m_MSBBit, "0000") + "(LSB)] = "
                            m_HexStr = auto_Value2HexStr("b" + m_bitstrL, m_bitwidth)
                        End If
                        m_decimal = m_HexStr
                    Else
                        If (m_bitwidth <= 31) Then
                            m_tmpStr = Replace(UCase(CStr(m_HexStr)), "0X", "", 1, 1)
                            m_decimal = CDbl("&H" & m_tmpStr) ''''Here it's Hex2Dec
                        Else
                            m_decimal = m_HexStr
                        End If
                    End If
        
                    m_tmpStr = FormatNumeric(" [" + m_bitstrL + "] ", -60) + m_HexStr
    '                If (m_algorithm = "crc") Then m_tmpStr = FormatNumeric(" [" + m_bitstrM + "] ", -40) + m_hexStr
                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 30) + m_tmpStr
                    TheExec.Datalog.WriteComment tmpdlgStr
                End If
            Next i
        Next m_Site

    ElseIf (FuseType = eFuse_UID) Then
            For Each m_Site In TheExec.sites
            m_pgmBitArr = InWave(m_Site).Data
            TheExec.Datalog.WriteComment ""
            For i = 0 To UBound(UIDFuse.Category)
                With UIDFuse.Category(i)
                    m_catename = .Name
                    m_stage = LCase(.Stage)
                    m_algorithm = LCase(.algorithm)
                    m_MSBBit = .MSBbit
                    m_LSBbit = .LSBbit
                    m_bitwidth = .BitWidth
                    m_defval = .DefaultValue
                    'm_defreal = LCase(.Default_Real)
                End With
                
                ''''only display the programming category
                If (m_stage = gS_JobName) Then

                    ''''PgmBit datalog format
                    tmpdlgStr = "Site(" + CStr(m_Site) + ") Programming : " + FormatNumeric(m_catename, gI_CFG_catename_maxLen)
                    tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "(LSB)] = "

                    m_bitStrM = ""
                    m_bitStrM = UIDFuse.Category(i).Write.BitStrM
                    m_bitstrL = UIDFuse.Category(i).Write.BitStrL
                    m_HexStr = UIDFuse.Category(i).Write.HexStr
                    m_tmpbitStrM = " [" + m_bitStrM + "] "
                    m_tmpStr = FormatNumeric(m_tmpbitStrM, -40) + m_HexStr
                                        
                    If (m_algorithm = "uid") Then
                    
                    ElseIf (m_algorithm = "crc") Then
                        For j = 0 To m_bitwidth - 1
                            bcnt = m_LSBbit + j
                            m_bitStrM = CStr(m_pgmBitArr(bcnt)) + m_bitStrM
                        Next j
                        m_bitstrL = StrReverse(m_bitStrM)
                        m_HexStr = auto_Value2HexStr("b" + m_bitStrM, m_bitwidth)
    
                        m_tmpbitStrM = " [" + m_bitStrM + "] "
                        m_tmpStr = FormatNumeric(m_tmpbitStrM, -40) + m_HexStr
                        If (m_bitwidth >= 32) Then m_decimal = m_HexStr
                        tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 12) + m_tmpStr
                        TheExec.Datalog.WriteComment tmpdlgStr
                    Else
                        If (m_bitwidth <= 31) Then
                            m_tmpStr1 = Replace(UCase(CStr(m_HexStr)), "0X", "", 1, 1)
                            m_decimal = CDbl("&H" & m_tmpStr1) ''''Here it's Hex2Dec
                        Else
                            m_decimal = m_HexStr
                        End If
                        tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 12) + m_tmpStr
                        TheExec.Datalog.WriteComment tmpdlgStr
                    End If
                End If ''''end of If (m_stage = gS_JobName) Then Next
            Next i
        Next m_Site
    Else
        If (FuseType = eFuse_CFG) Then
            m_Fuseblock = CFGFuse
        ElseIf (FuseType = eFuse_SEN) Then
            m_Fuseblock = SENFuse
        ElseIf (FuseType = eFuse_MON) Then
            m_Fuseblock = MONFuse
        ElseIf (FuseType = eFuse_UDR) Then
            m_Fuseblock = UDRFuse
        ElseIf (FuseType = eFuse_UDRE) Then
            m_Fuseblock = UDRE_Fuse
        ElseIf (FuseType = eFuse_UDRP) Then
            m_Fuseblock = UDRP_Fuse
        End If
        
        For Each m_Site In TheExec.sites
            m_pgmBitArr = InWave(m_Site).Data
            TheExec.Datalog.WriteComment ""
            For i = 0 To UBound(m_Fuseblock.Category)
                With m_Fuseblock.Category(i)
                    m_catename = .Name
                    m_stage = LCase(.Stage)
                    m_algorithm = LCase(.algorithm)
                    m_MSBBit = .MSBbit
                    m_LSBbit = .LSBbit
                    m_bitwidth = .BitWidth
                    m_defval = .DefaultValue
                    m_defreal = LCase(.Default_Real)
                    m_resolution = .Resoultion
                End With
                
                ''''only display the programming category
                If (m_stage = gS_JobName) Then

                    ''''PgmBit datalog format
                    tmpdlgStr = "Site(" + CStr(m_Site) + ") Programming : " + FormatNumeric(m_catename, gI_CFG_catename_maxLen)
                    tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "(LSB)] = "

                    m_bitStrM = ""
                    For j = 0 To m_bitwidth - 1
                        bcnt = m_LSBbit + j
                        m_bitStrM = CStr(m_pgmBitArr(bcnt)) + m_bitStrM
                    Next j
                    m_bitstrL = StrReverse(m_bitStrM)
                    m_HexStr = auto_Value2HexStr("b" + m_bitStrM, m_bitwidth)

                    m_tmpbitStrM = " [" + m_bitStrM + "] "
                    m_tmpStr = FormatNumeric(m_tmpbitStrM, -40) + m_HexStr
                                        
'                    If (m_algorithm = "lotid") Then
'    ''''                    m_pgmlotidStr = ""
'    ''''                    For j = 0 To EcidCharPerLotId - 1
'    ''''                        m_tmpStr = ""
'    ''''                        m_tmpStr = Mid(m_bitStrM, 1 + j * EcidBitPerLotIdChar, EcidBitPerLotIdChar) ''''EcidBitPerLotIdChar=6
'    ''''                        m_pgmlotidStr = m_pgmlotidStr + auto_MappingBinStrtoChar(m_tmpStr)
'    ''''                    Next j
'    ''''                    m_decimal = m_pgmlotidStr
'
'                    Else
                        If (m_bitwidth <= 31) Then
                            m_tmpStr1 = Replace(UCase(CStr(m_HexStr)), "0X", "", 1, 1)
                            m_value = CDbl("&H" & m_tmpStr1) ''''Here it's Hex2Dec
                        Else
                            m_value = m_HexStr
                        End If
                        
                        If (m_algorithm = "ids") Then
                            m_value = m_value * m_resolution
                            m_value = Format(m_value, "####.0000") + "mA "
                            m_tmpStr = FormatNumeric(m_tmpbitStrM, -40) + m_HexStr
                            'm_tmpStr = FormatNumeric(m_tmpbitStrM, -20) + FormatNumeric(m_value, 20) + m_hexStr
                        'ElseIf (m_algorithm = "vddbin" And (m_defreal Like "safe*voltage" Or m_defreal = "bincut")) Then
                        ElseIf (m_algorithm = "vddbin") Then
                            m_value = gD_BaseVoltage + m_value * m_resolution
                            m_value = Format(m_value, "####.0000") + "mV "
                            m_tmpStr = FormatNumeric(m_tmpbitStrM, -40) + m_HexStr
                            'm_tmpStr = FormatNumeric(m_tmpbitStrM, -20) + FormatNumeric(m_value, 20) + m_hexStr
                        'ElseIf (m_algorithm = "base" And m_defreal Like "safe*voltage") Then ''''for vddbin and base
                        ElseIf (m_algorithm = "base") Then ''''for vddbin and base
                            m_value = (m_value + 1) * m_resolution
                            m_value = Format(m_value, "####.0000") + "mV "
                            m_tmpStr = FormatNumeric(m_tmpbitStrM, -40) + m_HexStr
                            'm_tmpStr = FormatNumeric(m_tmpbitStrM, -20) + FormatNumeric(m_value, 20) + m_hexStr
                        End If
'                    End If
                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_value, 12) + m_tmpStr
                    TheExec.Datalog.WriteComment tmpdlgStr
                End If ''''end of If (m_stage = gS_JobName) Then Next
            Next i
        Next m_Site
    End If
    
    TheExec.Datalog.WriteComment ""
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''201810XX New, update with Optional StartBit,
'Public Function auto_eFuse_print_PgmBitsWave_BitMap(ByVal pgmWave As DSPWave, Optional showPrint As Boolean = False, Optional startbit As Long = 0)
Public Function auto_eFuse_print_PgmBitsWave_BitMap(ByVal pgmWave As DSPWave, _
                                                    ByVal SampleSize As Long, _
                                                    ByVal ReadCycle As Long, _
                                                    Optional showPrint As Boolean = False, _
                                                    Optional StartBit As Long = 0)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_print_PgmBitsWave_BitMap"
    
    If (showPrint = False) Then Exit Function
    
    Dim m_Site As Variant
    Dim PgmBits_Arr() As Long
    
    ''Print out All program bits
    For Each m_Site In TheExec.sites
        PgmBits_Arr = pgmWave.Data
        Call auto_PrintAllPgmBits(PgmBits_Arr, ReadCycle, SampleSize, gDL_BitsPerCycle, StartBit)
        'Call auto_PrintAllPgmBits(PgmBits_Arr, gDL_ReadCycles, gDL_TotalBits, gDL_BitsPerCycle, StartBit)
    Next m_Site

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
      
End Function

''''201810XX New
Public Function auto_eFuse_print_DSSCReadWave_BitMap(ByVal FuseType As eFuseBlockType, Optional showPrint As Boolean = False)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_print_DSSCReadWave_BitMap"
    
    If (showPrint = False) Then Exit Function
    
    Dim m_Site As Variant
    Dim m_readBits_Arr() As Long
    Dim m_SingleBitWave As New DSPWave
    'Dim m_PrintRow As Long
    
    Select Case FuseType
        Case eFuse_ECID:
            Call rundsp.eFuse_DspWave_Copy(gDW_ECID_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_CFG:
            Call rundsp.eFuse_DspWave_Copy(gDW_CFG_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_UID:
            Call rundsp.eFuse_DspWave_Copy(gDW_UID_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_MON:
            Call rundsp.eFuse_DspWave_Copy(gDW_MON_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_SEN:
            Call rundsp.eFuse_DspWave_Copy(gDW_SEN_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_UDR:
            Call rundsp.eFuse_DspWave_Copy(gDW_UDR_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_UDRE:
            Call rundsp.eFuse_DspWave_Copy(gDW_UDRE_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_UDRP:
            Call rundsp.eFuse_DspWave_Copy(gDW_UDRP_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_CMP:
            Call rundsp.eFuse_DspWave_Copy(gDW_CMP_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_CMPE:
            Call rundsp.eFuse_DspWave_Copy(gDW_CMPE_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_CMPP:
            Call rundsp.eFuse_DspWave_Copy(gDW_CMPP_Read_SingleBitWave, m_SingleBitWave)
        Case Else:
            GoTo errHandler
    End Select
    For Each m_Site In TheExec.sites
        m_readBits_Arr = m_SingleBitWave(m_Site).Data
        Call auto_PrintBitMap(m_readBits_Arr, gDL_ReadCycles, gDL_TotalBits, gDL_BitsPerCycle, FuseType)
    Next m_Site
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
      
End Function

''''201808XX New, Here inWave is "doubleBitsWave"
''Public Function auto_eFuse_print_DSSCReadWave_Category(ByVal fusetype As eFuseBlockType, ByVal inWave As DSPWave, _
''                                                       Optional byStage As Boolean = False, Optional showPrint As Boolean = True) As Long

Public Function auto_eFuse_print_DSSCReadWave_Category(ByVal FuseType As eFuseBlockType, _
                                                       Optional byStage As Boolean = False, Optional showPrint As Boolean = False, _
                                                       Optional PatName As String = "") As Long
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_print_DSSCReadWave_Category"

    If (showPrint = False) Then Exit Function
    
    Dim i As Long, j As Long, k As Long, bcnt As Long
    Dim m_Site As Variant
    Dim m_readBitArr() As Long
    Dim m_tmpStr As String
    Dim m_stage As String
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_MSBBit As Long
    Dim m_LSBbit As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant
    Dim m_value As Variant
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim tmpdlgStr As String
    Dim m_bitstrL As String
    Dim m_bitStrM As String
    Dim m_tmpVal As Variant
    Dim m_HexStr As String
    Dim m_readlotidStr As String
    Dim m_namelen As Long
    Dim m_condcnt As Long
    Dim m_read_pkgname As String
    Dim m_condHeadStr As String
    Dim m_condhexStr As String
    Dim m_tmpbitStrM As String
    Dim mStr_BankName As String:: mStr_BankName = ""
    Dim m_CalcSum As Long

    Dim m_Fuseblock As EFuseCategorySyntax

    Select Case FuseType
        Case eFuse_ECID:
            m_Fuseblock = ECIDFuse
            m_namelen = gI_ECID_catename_maxLen

        Case eFuse_CFG:
            m_Fuseblock = CFGFuse
            m_namelen = gI_CFG_catename_maxLen

        Case eFuse_UID:
            m_Fuseblock = UIDFuse
            m_namelen = gI_UID_catename_maxLen
        Case eFuse_MON:
            m_Fuseblock = MONFuse
            m_namelen = gI_MON_catename_maxLen
        Case eFuse_SEN:
            m_Fuseblock = SENFuse
            m_namelen = gI_SEN_catename_maxLen
        Case eFuse_UDR:
            m_Fuseblock = UDRFuse
            m_namelen = gI_UDR_catename_maxLen
            mStr_BankName = "UDR"
        Case eFuse_UDRE:
            m_Fuseblock = UDRE_Fuse
            m_namelen = gI_UDRE_catename_maxLen
            mStr_BankName = "UDRE"
        Case eFuse_UDRP:
            m_Fuseblock = UDRP_Fuse
            m_namelen = gI_UDRP_catename_maxLen
            mStr_BankName = "UDRP"
        Case eFuse_CMP:
            m_Fuseblock = CMPFuse
            m_namelen = gI_CMP_catename_maxLen
            mStr_BankName = "CMP"
        Case eFuse_CMPE:
            m_Fuseblock = CMPE_Fuse
            m_namelen = gI_CMPE_catename_maxLen
            mStr_BankName = "CMPE"
        Case eFuse_CMPP:
            m_Fuseblock = CMPP_Fuse
            m_namelen = gI_CMPP_catename_maxLen
            mStr_BankName = "CMPP"
        Case Else:
            GoTo errHandler
    End Select

    Dim m_match_flag As Boolean
    Dim m_UboundSize As Long:: m_UboundSize = UBound(m_Fuseblock.Category)
    If (FuseType = eFuse_ECID) Then m_UboundSize = m_UboundSize - 1
    
    ''''<NOTICE> Will update for case eFuse_ECID later on.
    
        For Each m_Site In TheExec.sites
            If (FuseType = eFuse_UDR Or FuseType = eFuse_UDRE Or FuseType = eFuse_UDRP Or FuseType = eFuse_CMP Or FuseType = eFuse_CMPE Or FuseType = eFuse_CMPP) Then
                TheExec.Datalog.WriteComment ""
                tmpdlgStr = "Site(" + CStr(m_Site) + ") , " + mStr_BankName + " , pat: " + PatName
                TheExec.Datalog.WriteComment tmpdlgStr
                tmpdlgStr = ""
            End If
            m_condcnt = 0
            TheExec.Datalog.WriteComment ""
            For i = 0 To m_UboundSize
            'For i = 0 To UBound(m_Fuseblock.Category)
                m_stage = LCase(m_Fuseblock.Category(i).Stage)
                m_match_flag = False
                
                If (byStage = True And m_stage = gS_JobName) Then
                    m_match_flag = True
                ElseIf (byStage = False) Then
                    m_match_flag = True
                End If
                
                If (m_match_flag = True) Then
                    With m_Fuseblock.Category(i)
                        m_catename = .Name
                        ''m_stage = LCase(.Stage)
                        m_algorithm = LCase(.algorithm)
                        m_MSBBit = .MSBbit
                        m_LSBbit = .LSBbit
                        m_bitwidth = .BitWidth
                        m_defval = .DefaultValue
                        m_defreal = LCase(.Default_Real)
                        ''''Read elements
                        m_decimal = .Read.Decimal
                        m_bitStrM = .Read.BitStrM
                        m_bitstrL = .Read.BitStrL
                        m_HexStr = .Read.HexStr
                        m_value = .Read.Value
                    End With
        
'                    ''''ReadBit datalog format
'                    tmpdlgStr = "Site(" + CStr(m_Site) + ") Read from DSSC : " + FormatNumeric(m_catename, m_namelen)
'                    tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_msbbit, "0000") + ":" + Format(m_lsbbit, "0000") + "(LSB)] = "
'                    m_tmpbitStrM = " [" + m_bitstrM + "] "
'                    m_tmpStr = FormatNumeric(m_tmpbitStrM, -40) + m_hexStr
'                    If (m_lsbbit > m_msbbit) Then
'                        m_bitstrM = m_bitstrL
'                    End If
                    
                    If (m_algorithm = "lotid") Then
                        m_readlotidStr = ""
                        If (m_LSBbit > m_MSBBit) Then
                             Dim tmp As String
                            tmp = m_bitStrM
                            m_bitStrM = m_bitstrL
                        End If
                        For j = 0 To EcidCharPerLotId - 1
                            m_tmpStr = ""
                            m_tmpStr = Mid(m_bitStrM, 1 + j * EcidBitPerLotIdChar, EcidBitPerLotIdChar) ''''EcidBitPerLotIdChar=6
                            m_readlotidStr = m_readlotidStr + auto_MappingBinStrtoChar(m_tmpStr)
                        Next j
                        m_value = m_readlotidStr
                        If (m_LSBbit > m_MSBBit) Then
                            m_bitStrM = tmp
                        End If
                        ''m_tmpStr = FormatNumeric(m_tmpbitStrM, -40) + m_hexStr

                    ElseIf (m_algorithm = "cond" And m_condcnt = 0) Then
                        m_condcnt = m_condcnt + 1
                        m_condhexStr = auto_Value2HexStr("b" + gS_CFG_Cond_Read_bitStrM)
                        m_condHeadStr = "Site(" + CStr(m_Site) + ") Read from DSSC : " + FormatNumeric("CFG_Condition", m_namelen)
                        m_condHeadStr = FormatNumeric(m_condHeadStr, -30) + "                      " + " =" + FormatNumeric(gS_CFG_Cond_Read_pkgname, 15) + "  " + m_condhexStr
                        'm_condHeadStr = m_condHeadStr + " " + gS_CFG_Cond_Read_pkgname + " = " + m_condhexStr
                        TheExec.Datalog.WriteComment m_condHeadStr
                    
                    ElseIf (m_algorithm = "ids") Then
                        m_value = Format(m_value, "####0.0000") + "mA "
                        'm_tmpStr = FormatNumeric(m_tmpbitStrM, -20) + FormatNumeric(m_value, 20) + m_HexStr
                    ElseIf (m_algorithm = "vddbin" Or m_algorithm = "base") Then
                    'ElseIf (m_defreal Like "safe*voltage" Or m_defreal = "bincut") Then ''''for vddbin and base
                        m_value = Format(m_value, "####0.0000") + "mV "
                        ''20191202 For Central Meeting
                        ''if the field of bincut and safe voltage isn't fused, then the printing will show "0.0000mV".
                        m_CalcSum = m_Fuseblock.Category(i).Read.BitArrWave.CalcSum
                        If (m_CalcSum = 0) Then m_value = Format(0, "####0.0000") + "mV "
                        m_tmpStr = FormatNumeric(m_tmpbitStrM, -20) + FormatNumeric(m_value, 20) + m_HexStr
                    ElseIf (m_defreal Like "decimal") Then ''''for base with decimal
                        m_CalcSum = m_Fuseblock.Category(i).Read.BitArrWave.CalcSum
                        If (m_CalcSum = 0) Then m_value = Format(0, "####0.0000") + "mV "
                    ElseIf (m_algorithm = "crc") Then
                        ''''use hexStr to present the data
                        m_value = m_HexStr
                    Else
                        ''''nothing for now, add something if anything.
                    End If
                    
                                        ''''ReadBit datalog format
                    tmpdlgStr = "Site(" + CStr(m_Site) + ") Read from DSSC : " + FormatNumeric(m_catename, m_namelen)
                    tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "(LSB)] = "
                    m_tmpbitStrM = " [" + m_bitStrM + "] "
                    m_tmpStr = FormatNumeric(m_tmpbitStrM, -40) + m_HexStr
                    
                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_value, 12) + m_tmpStr
                    TheExec.Datalog.WriteComment tmpdlgStr
                End If
            Next i
        Next m_Site


''''-------------------------------------------------------------------
''    If (fusetype = eFuse_ECID) Then
''        For Each m_Site In TheExec.Sites
''            m_readBitArr = inWave(m_Site).Data
''            TheExec.Datalog.WriteComment ""
''            For i = 0 To UBound(ECIDFuse.Category) - 1
''                With ECIDFuse.Category(i)
''                    m_catename = .Name
''                    m_stage = LCase(.Stage)
''                    m_algorithm = LCase(.Algorithm)
''                    m_msbbit = .MSBbit
''                    m_lsbbit = .LSBbit
''                    m_bitwidth = .Bitwidth
''                    m_defval = .DefaultValue
''                    m_defreal = LCase(.Default_Real)
''                End With
''
''                ''''ReadBit datalog format
''                tmpdlgStr = "Site(" + CStr(m_Site) + ") Read from DSSC : " + FormatNumeric(m_catename, gI_ECID_catename_maxLen)
''                If (UCase(ECIDFuse.Category(i).MSBFirst) = "Y") Then
''                    If (m_algorithm = "crc") Then
''                        ''''<NOTICE>20170331, ECID CRC always coding/fusing as [MSB......LSB]
''                        ''''Example: m_LSBbit(255) is MSB of CRC result, m_MSBbit(240) is LSB of CRC result
''                        ''''         gL_ECID_CRC_MSB = m_LSBbit(255), gL_ECID_CRC_LSB = m_MSBbit(240)
''                        ''''The above is correct in the function auto_ECIDConstant_Initialize()
''                        tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(gL_ECID_CRC_MSB, "0000") + ":" + Format(gL_ECID_CRC_LSB, "0000") + "(LSB)] = "
''                        ''m_tmpStr = " [" + m_bitStrM + "]"
''                    Else
''                        tmpdlgStr = tmpdlgStr + " [(LSB)" + Format(m_lsbbit, "0000") + ":" + Format(m_msbbit, "0000") + "(MSB)] = "
''                        ''m_tmpStr = " [" + m_bitStrL + "]"
''                    End If
''                Else
''                    tmpdlgStr = tmpdlgStr + " [(MSB)" + Format(m_msbbit, "0000") + ":" + Format(m_lsbbit, "0000") + "(LSB)] = "
''                    ''m_tmpStr = " [" + m_bitStrM + "]"
''                End If
''
''
''                m_bitStrM = ""
''                For j = 0 To m_bitwidth - 1
''                    bcnt = m_msbbit + j
''                    m_bitStrM = m_bitStrM + CStr(m_readBitArr(bcnt))
''                Next j
''                m_bitStrL = StrReverse(m_bitStrM)
''                m_hexStr = auto_Value2HexStr("b" + m_bitStrM, m_bitwidth)
''
''                If (m_algorithm = "lotid") Then
''                    m_readlotidStr = ""
''                    For j = 0 To EcidCharPerLotId - 1
''                        m_tmpStr = ""
''                        m_tmpStr = Mid(m_bitStrM, 1 + j * EcidBitPerLotIdChar, EcidBitPerLotIdChar) ''''EcidBitPerLotIdChar=6
''                        m_readlotidStr = m_readlotidStr + auto_MappingBinStrtoChar(m_tmpStr)
''                    Next j
''                    m_decimal = m_readlotidStr
''                Else
''                    If (m_bitwidth <= 31) Then
''                        m_tmpStr = Replace(UCase(CStr(m_hexStr)), "0X", "", 1, 1)
''                        m_decimal = CDbl("&H" & m_tmpStr) ''''Here it's Hex2Dec
''                    Else
''                        m_decimal = m_hexStr
''                    End If
''                End If
''
''                m_tmpStr = FormatNumeric(" [" + m_bitStrL + "] ", -40) + m_hexStr
''                tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 12) + m_tmpStr
''                TheExec.Datalog.WriteComment tmpdlgStr
''            Next i
''        Next m_Site
''    End If

    TheExec.Datalog.WriteComment ""
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function



''''201812XX Update
Public Function auto_eFuse_setReadData(ByVal FuseType As eFuseBlockType) As Boolean
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_setReadData"
   
    Dim m_Site As Variant
    Dim i As Long, j As Long
    Dim m_catename As String
    Dim m_startbit As Long
    Dim m_stopbit As Long
    Dim m_bitwidth As Long
    Dim m_stage As String
    Dim m_algorithm As String
    Dim m_dec As New SiteVariant
    Dim m_bitsum As New SiteLong
    Dim m_bitWave As New DSPWave

    Dim m_readDecArr() As Double
    Dim m_dblBitArr() As Long
    Dim m_bitArr() As Long
    Dim m_FuseRead As EFuseCategoryParamResultSyntax
    Dim m_Fuseblock As EFuseCategorySyntax
    Dim m_bitStrM As String
    Dim m_bitstrL As String
    Dim m_HexStr As String
    Dim m_resolution As Double
    
''----New
    Dim m_dfreal As String
    Dim m_calc_mode As Long
    Dim mSD_dec As New SiteDouble
    Dim mSD_val As New SiteDouble
    Dim mSL_bitSum As New SiteLong
    Dim mSV_bitStrM As New SiteVariant
    Dim mSV_bitStrL As New SiteVariant
    Dim mSV_hexStr As New SiteVariant
    Dim mSB_cmpResult As New SiteBoolean
    Dim mSV_decimal As New SiteVariant
    Dim mSV_value As New SiteVariant
''---

    If (FuseType = eFuse_ECID) Then
        m_Fuseblock = ECIDFuse
    ElseIf (FuseType = eFuse_CFG) Then
        m_Fuseblock = CFGFuse
    ElseIf (FuseType = eFuse_UID) Then
    ElseIf (FuseType = eFuse_SEN) Then
    ElseIf (FuseType = eFuse_MON) Then
    ElseIf (FuseType = eFuse_UDR) Then
    End If

    If (FuseType = eFuse_ECID) Then
''''        For Each m_Site In TheExec.Sites
''''            m_readDecArr = readDecCateWave.Data
''''            For i = 0 To UBound(m_FuseBlock.Category)
''''                With m_FuseBlock.Category(i)
''''                    m_catename = .Name
''''                    m_stage = LCase(.Stage)
''''                    m_startbit = .MSBbit ''''<NOTICE>
''''                    m_stopbit = .LSBbit
''''                    m_bitwidth = .bitwidth
''''                    m_algorithm = LCase(.Algorithm)
''''                    m_resolution = .Resoultion
''''                End With
''''                m_dec = m_readDecArr(i)
''''                m_bitWave = doubleBitWave.Select(m_startbit, 1, m_bitwidth).Copy
''''                m_bitSum = m_bitWave.CalcSum
''''                m_bitArr = m_bitWave.Data
''''
''''                m_bitstrM = ""
''''                For j = 0 To m_bitwidth - 1
''''                    m_bitstrM = m_bitstrM + CStr(m_bitArr(j))
''''                Next j
''''                m_bitStrL = StrReverse(m_bitstrM)
''''                m_hexStr = auto_Value2HexStr("b" + m_bitstrM, m_bitwidth)
''''
''''                If (m_dec = -9999) Then
''''                   m_dec = m_hexStr
''''                End If
''''                ''''Here is just to get Read.Decimal/.BitSummation/.bitArrWave, to save the TTR
''''                ''''for other properties of Read, they are leaved in the print section if needed.
''''                With m_FuseBlock.Category(i).Read
''''                    .Decimal = m_dec
''''                    .BitSummation = m_bitSum
''''                    .BitArrWave = m_bitWave.Copy
''''                    .hexStr = m_hexStr
''''                    .bitstrM = m_bitstrM
''''                    .BitStrL = m_bitStrL
''''                    .Value = m_dec
''''                    .ValStr = CStr(m_dec)
''''                End With
''''            Next i
''''        Next m_Site

    Else

        gS_CFG_Cond_Read_bitStrM = "" ''''<MUST>
        For i = 0 To UBound(m_Fuseblock.Category)
            With m_Fuseblock.Category(i)
                m_catename = .Name
                m_stage = LCase(.Stage)
                m_startbit = .LSBbit ''''<NOTICE>
                m_stopbit = .MSBbit
                m_bitwidth = .BitWidth
                m_algorithm = LCase(.algorithm)
                m_resolution = .Resoultion
                m_dfreal = LCase(.Default_Real)
            End With
            
            m_calc_mode = 0 ''''default
            ''''If (m_algorithm = "vddbin"/"base" And m_dfreal = "decimal") Then m_vddbin_mode = 0
            If (m_algorithm = "vddbin" And m_dfreal Like "safe*voltage") Then
                m_calc_mode = 1
            ElseIf (m_algorithm = "vddbin" And m_dfreal = "bincut") Then
                m_calc_mode = 2
            ElseIf (m_algorithm = "base" And m_dfreal Like "safe*voltage") Then
                m_calc_mode = 3
            End If

            ''''<NOTICE> m_bitWave, its Element(0) is always LSBbit value
            Call rundsp.eFuse_Get_ValueFromWave(FuseType, i, m_resolution, m_calc_mode, mSD_dec, mSD_val, mSL_bitSum, m_bitWave)

            mSV_decimal = mSD_dec
            mSV_value = mSD_val
            For Each m_Site In TheExec.sites
                m_bitArr = m_bitWave.Data
                mSV_bitStrM = ""
                For j = 0 To m_bitwidth - 1
                    mSV_bitStrM = CStr(m_bitArr(j)) + mSV_bitStrM
                Next j
                mSV_bitStrL = StrReverse(mSV_bitStrM)
                mSV_hexStr = auto_Value2HexStr("b" + mSV_bitStrM, m_bitwidth)

                ''''For the CFG Condition Bits (Read)
                If (m_algorithm = "cond" And FuseType = eFuse_CFG) Then
                    gS_CFG_Cond_Read_bitStrM = mSV_bitStrM + gS_CFG_Cond_Read_bitStrM
                End If

                If (mSD_dec = -9999) Then
                    mSV_decimal = mSV_hexStr
                    mSV_value = mSV_hexStr
                End If
            Next m_Site


            ''''Here is just to get Read.Decimal/.BitSummation/.bitArrWave, to save the TTR
            ''''for other properties of Read, they are leaved in the print section if needed.
            With m_Fuseblock.Category(i).Read
                .Decimal = mSV_decimal
                .BitSummation = mSL_bitSum
                For Each m_Site In TheExec.sites
                    .BitArrWave = m_bitWave.Copy
                Next m_Site
                .HexStr = mSV_hexStr
                .BitStrM = mSV_bitStrM
                .BitStrL = mSV_bitStrL
                .Value = mSV_value
                .ValStr = .Value
            End With
            
            If (i = UBound(m_Fuseblock.Category) And FuseType = eFuse_CFG) Then
                gS_CFG_Cond_Read_pkgname = "NA"
                gL_CFG_Cond_compResult = -99 '''Fail
                mSB_cmpResult = gS_CFG_Cond_Read_bitStrM.compare(EqualTo, gS_CFGCondTable_bitsStr)
                ''''Boolean True=-1, False=0
                gL_CFG_Cond_compResult = mSB_cmpResult
                gL_CFG_Cond_compResult = gL_CFG_Cond_compResult.Add(1) ''''pass:0, fail:1
                For Each m_Site In TheExec.sites
                    If (mSB_cmpResult = True) Then
                        gS_CFG_Cond_Read_pkgname = UCase(gS_cfgFlagname)
                    End If
                Next m_Site
            End If
        Next i

''''------------------
'If (False) Then
'        For Each m_Site In TheExec.Sites
'            gS_CFG_Cond_Read_bitStrM(m_Site) = ""
'            m_readDecArr = readDecCateWave.Data
'            ''m_dblBitArr = doubleBitWave.Data
'            For i = 0 To UBound(m_FuseBlock.Category)
'                With m_FuseBlock.Category(i)
'                    m_catename = .Name
'                    m_stage = LCase(.Stage)
'                    m_startbit = .LSBbit ''''<NOTICE>
'                    m_stopbit = .MSBbit
'                    m_bitwidth = .Bitwidth
'                    m_algorithm = LCase(.Algorithm)
'                    m_resolution = .Resoultion
'                End With
'                m_dec = m_readDecArr(i)
'                m_bitWave = doubleBitWave.Select(m_startbit, 1, m_bitwidth).Copy
'                m_bitSum = m_bitWave.CalcSum
'                m_bitArr = m_bitWave.Data
'
'                m_bitStrM = ""
'                For j = 0 To m_bitwidth - 1
'                    m_bitStrM = CStr(m_bitArr(j)) + m_bitStrM
'                Next j
'                m_bitStrL = StrReverse(m_bitStrM)
'                m_hexStr = auto_Value2HexStr("b" + m_bitStrM, m_bitwidth)
'
'                ''''For the CFG Condition Bits (Read)
'                If (m_algorithm = "cond" And fusetype = eFuse_CFG) Then
'                    gS_CFG_Cond_Read_bitStrM = m_bitStrM + gS_CFG_Cond_Read_bitStrM
'
'                ElseIf (m_algorithm = "ids") Then
'                ElseIf (m_algorithm = "vddbin") Then
'                End If
'
'                If (m_dec = -9999) Then
'                   m_dec = m_hexStr
'                End If
'                ''''Here is just to get Read.Decimal/.BitSummation/.bitArrWave, to save the TTR
'                ''''for other properties of Read, they are leaved in the print section if needed.
'                With m_FuseBlock.Category(i).Read
'                    .Decimal = m_dec
'                    .BitSummation = m_bitSum
'                    .bitArrWave = m_bitWave.Copy
'                    .HexStr = m_hexStr
'                    .BitStrM = m_bitStrM
'                    .BitStrL = m_bitStrL
'                    .Value = m_dec
'                    .ValStr = CStr(m_dec)
'                End With
'            Next i
'
'            ''''201812XX update
'            If (fusetype = eFuse_CFG) Then
'                gS_CFG_Cond_Read_pkgname = "NA"
'                gL_CFG_Cond_compResult = -1 '''Fail
'                If (gS_CFG_Cond_Read_bitStrM = gS_CFGCondTable_bitsStr) Then
'                    gS_CFG_Cond_Read_pkgname = UCase(gS_cfgFlagname)
'                    gL_CFG_Cond_compResult = 0 ''''Pass
'                End If
'            End If
'        Next m_Site
'End If ''''end of If (False) Then
''''------------------

    End If
    
    ''''<MUST>
    If (FuseType = eFuse_ECID) Then
        ECIDFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_CFG) Then
        CFGFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_UID) Then
    ElseIf (FuseType = eFuse_SEN) Then
    ElseIf (FuseType = eFuse_MON) Then
    ElseIf (FuseType = eFuse_UDR) Then
    End If

    ''''debug purpose to check if the data are set to Read structure.
    If (False) Then
        For i = 0 To UBound(m_Fuseblock.Category)
            For Each m_Site In TheExec.sites
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.Decimal = " & m_Fuseblock.Category(i).Read.Decimal
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.Value   = " & m_Fuseblock.Category(i).Read.Value
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.HexStr  = " & m_Fuseblock.Category(i).Read.HexStr
            Next m_Site
        Next i
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'auto_eFuse_setReadData_forDatalog
'auto_eFuse_setReadData_forSyntax
''''201901XX Update
Public Function auto_eFuse_setReadData_forSyntax(ByVal FuseType As eFuseBlockType) As Boolean
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_setReadData_forSyntax"
   
    Dim m_Site As Variant
    Dim i As Long, j As Long
    Dim m_catename As String
    Dim m_startbit As Long
    Dim m_stopbit As Long
    Dim m_bitwidth As Long
    Dim m_stage As String
    Dim m_algorithm As String
''    Dim m_dec As New SiteVariant
''    Dim m_bitSum As New SiteLong
    Dim m_bitWave As New DSPWave

    Dim m_readDecArr() As Double
    Dim m_dblBitArr() As Long
    Dim m_bitArr() As Long
    Dim m_FuseRead As EFuseCategoryParamResultSyntax
    Dim m_Fuseblock As EFuseCategorySyntax
'    Dim m_bitstrM As String
'    Dim m_bitStrL As String
'    Dim m_hexStr As String
    Dim m_resolution As Double
    
''----New
    Dim m_dfreal As String
    Dim m_calc_mode As Long
    Dim mSD_dec As New SiteDouble
    Dim mSD_val As New SiteDouble
    Dim mSL_bitSum As New SiteLong
    Dim mSV_bitStrM As New SiteVariant
    Dim mSV_bitStrL As New SiteVariant
    Dim mSV_hexStr As New SiteVariant
    Dim mSB_cmpResult As New SiteBoolean
    Dim mSV_decimal As New SiteVariant
    Dim mSV_value As New SiteVariant
    
    Dim m_readDecCateWave As New DSPWave
    Dim mSD_readDecArr() As New SiteDouble
    Dim m_DoubleBitWave As New DSPWave
    Dim m_doubleBitArr() As Long
''---

    If (FuseType = eFuse_ECID) Then
        m_Fuseblock = ECIDFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_ECID_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_ECID_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_CFG) Then
        m_Fuseblock = CFGFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_CFG_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_CFG_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_UID) Then
        m_Fuseblock = UIDFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_UID_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_UID_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_SEN) Then
        m_Fuseblock = SENFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_SEN_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_SEN_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_MON) Then
        m_Fuseblock = MONFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_MON_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_MON_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_UDR) Then
        m_Fuseblock = UDRFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_UDR_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_UDR_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_UDRE) Then
        m_Fuseblock = UDRE_Fuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_UDRE_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_UDRE_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_UDRP) Then
        m_Fuseblock = UDRP_Fuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_UDRP_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_UDRP_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_CMP) Then
        m_Fuseblock = CMPFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_CMP_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_CMP_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_CMPE) Then
        m_Fuseblock = CMPE_Fuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_CMPE_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_CMPE_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_CMPP) Then
        m_Fuseblock = CMPP_Fuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_CMPP_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_CMPP_Read_DoubleBitWave.Copy
        Next m_Site
    End If

    ReDim mSD_readDecArr(UBound(m_Fuseblock.Category))
    
    For Each m_Site In TheExec.sites
        m_readDecArr = m_readDecCateWave.Data
        For i = 0 To UBound(m_readDecArr)
            mSD_readDecArr(i) = m_readDecArr(i)
        Next i
    Next m_Site
    
    If (FuseType = eFuse_ECID) Then
    'Dim m_user_proberSite As New SiteVariant
    
        For i = 0 To UBound(m_Fuseblock.Category)
            With m_Fuseblock.Category(i)
                m_catename = .Name
                m_stage = LCase(.Stage)
                If (.LSBbit < .MSBbit) Then
                    m_startbit = .LSBbit ''''<NOTICE>
                    m_stopbit = .MSBbit
                Else
                    m_startbit = .MSBbit ''''<NOTICE>
                    m_stopbit = .LSBbit
                End If
                m_bitwidth = .BitWidth
                m_algorithm = LCase(.algorithm)
                m_resolution = .Resoultion
                m_dfreal = LCase(.Default_Real)
            End With
            
            mSV_decimal = mSD_readDecArr(i)
            mSV_value = mSV_decimal
            

            For Each m_Site In TheExec.sites
                m_bitWave = m_DoubleBitWave.Select(m_startbit, 1, m_bitwidth).Copy
                Call auto_eFuse_bitWave_to_binStr_HexStr(m_bitWave, mSV_bitStrM, mSV_hexStr, True, False)
                mSV_hexStr = auto_Value2HexStr("b" + mSV_bitStrM, m_bitwidth) ''''<MUST> for syntax comparison
            Next m_Site

            'mSV_value = mSV_hexStr

            If (m_algorithm = "lotid") Then
                Dim tmp As String
                Dim m_ReadLotID As String:: m_ReadLotID = ""
                For Each m_Site In TheExec.sites
                    mSV_bitStrL(m_Site) = StrReverse(mSV_bitStrM(m_Site))
                    For j = 0 To EcidCharPerLotId - 1
                        tmp = ""
                        tmp = Mid(mSV_bitStrL(m_Site), j * EcidCharPerLotId + 1, EcidCharPerLotId)
                        m_ReadLotID = m_ReadLotID + auto_MappingBinStrtoChar(tmp)

                    Next j
                    'mSV_value(m_Site) = auto_MappingBinStrtoChar(mSV_bitStrM(m_Site)) + m_ReadLotID
                    mSV_value(m_Site) = m_ReadLotID
                    m_ReadLotID = ""
                Next m_Site
                HramLotId = mSV_value
                'mSV_value = LotID
            ElseIf (LCase(m_catename) Like "*wafer*id") Then
                HramWaferId = mSV_decimal
            ElseIf (LCase(m_catename) Like "*x*coord*") Then
                HramXCoord = mSV_decimal
            ElseIf (LCase(m_catename) Like "*y*coord*") Then
                HramYCoord = mSV_decimal
            End If

            With m_Fuseblock.Category(i).Read
                .Decimal = mSV_decimal
                .Value = mSV_value
                .ValStr = mSV_value
                .HexStr = mSV_hexStr
            End With
        Next i

    Else
        'gS_CFG_Cond_Read_bitStrM = "" ''''<MUST>
        Set gS_CFG_Cond_Read_bitStrM = Nothing

        For i = 0 To UBound(m_Fuseblock.Category)
            With m_Fuseblock.Category(i)
                m_catename = .Name
                m_stage = LCase(.Stage)
                m_startbit = .LSBbit ''''<NOTICE>
                m_stopbit = .MSBbit
                m_bitwidth = .BitWidth
                m_algorithm = LCase(.algorithm)
                m_resolution = .Resoultion
                m_dfreal = LCase(.Default_Real)
            End With
            
            mSV_decimal = mSD_readDecArr(i)
            mSV_value = mSV_decimal
            If (m_resolution = 0#) Then m_resolution = 1#
            m_calc_mode = 0 ''''default

            ''''If (m_algorithm = "vddbin"/"base" And m_dfreal = "decimal") Then m_vddbin_mode = 0
            If (m_algorithm = "vddbin" And m_dfreal = "default") Then
            'If (m_algorithm = "vddbin" And m_dfreal Like "safe*voltage") Then
                m_calc_mode = 1
            ElseIf (m_algorithm = "vddbin" And m_dfreal = "bincut") Then
                m_calc_mode = 2
            ElseIf (m_algorithm = "base" And m_dfreal = "default") Then
            'ElseIf (m_algorithm = "base" And m_dfreal Like "safe*voltage") Then
                m_calc_mode = 3
            End If

            ''''For the CFG Condition Bits (Read)
            If (m_algorithm = "cond" And FuseType = eFuse_CFG) Then
                For Each m_Site In TheExec.sites
                    m_bitWave = m_DoubleBitWave.Select(m_startbit, 1, m_bitwidth).Copy
                    Call auto_eFuse_bitWave_to_binStr_HexStr(m_bitWave, mSV_bitStrM, mSV_hexStr, True, False)
                    gS_CFG_Cond_Read_bitStrM = mSV_bitStrM + gS_CFG_Cond_Read_bitStrM
                Next m_Site
                If (m_bitwidth >= 32) Then
                    mSV_decimal = mSV_hexStr
                    mSV_value = mSV_hexStr
                End If
            ''End If
            ElseIf (m_algorithm = "uid") Then
                For Each m_Site In TheExec.sites
                    mSV_decimal(m_Site) = m_DoubleBitWave(m_Site).Select(m_startbit, 1, m_bitwidth).CalcSum
                    mSV_value(m_Site) = mSV_decimal(m_Site) / m_bitwidth
                Next m_Site
                mSV_decimal = mSV_value
            ElseIf (m_bitwidth >= 32 Or mSV_decimal.compare(EqualTo, -9999).All(True) = True) Then
                For Each m_Site In TheExec.sites
                    m_bitWave = m_DoubleBitWave.Select(m_startbit, 1, m_bitwidth).Copy
                    Call auto_eFuse_bitWave_to_binStr_HexStr(m_bitWave, mSV_bitStrM, mSV_hexStr, True, False)
                    mSV_hexStr = auto_Value2HexStr("b" + mSV_bitStrM, m_bitwidth) ''''<MUST> for syntax comparison
                Next m_Site
                mSV_decimal = mSV_hexStr
                mSV_value = mSV_hexStr
            Else
                If (m_calc_mode = 0) Then
                    ''''ids: resolution<>0
                    ''''decimal: resolution=1
                    mSV_value = mSV_decimal.Multiply(m_resolution)
                ElseIf (m_calc_mode = 1) Then
                    ''''vddbin and safe voltage
                    mSV_value = mSV_decimal.Multiply(m_resolution).Add(gDD_BaseVoltage)
                ElseIf (m_calc_mode = 2) Then
                    ''''bincut, only limit is variant per dice
                    mSV_value = mSV_decimal.Multiply(m_resolution).Add(gDD_BaseVoltage)
                ElseIf (m_calc_mode = 3) Then
                    ''''base and safe voltage
                    mSV_value = mSV_decimal.Add(1).Multiply(m_resolution)
                End If
                For Each m_Site In TheExec.sites
                    mSV_hexStr = auto_Value2HexStr(mSV_decimal, m_bitwidth)
                Next m_Site
            End If
            
            With m_Fuseblock.Category(i).Read
                .Decimal = mSV_decimal
                .Value = mSV_value
                .ValStr = mSV_value
                .HexStr = mSV_hexStr
            End With

            If (i = UBound(m_Fuseblock.Category) And FuseType = eFuse_CFG) Then
                gS_CFG_Cond_Read_pkgname = "NA"
                gL_CFG_Cond_compResult = -99 '''Fail
                mSB_cmpResult = gS_CFG_Cond_Read_bitStrM.compare(EqualTo, gS_CFGCondTable_bitsStr)
                ''''Boolean True=-1, False=0
                gL_CFG_Cond_compResult = mSB_cmpResult
                gL_CFG_Cond_compResult = gL_CFG_Cond_compResult.Add(1) ''''pass:0, fail:1
                For Each m_Site In TheExec.sites
                    If (mSB_cmpResult = True) Then
                        gS_CFG_Cond_Read_pkgname = UCase(gS_cfgFlagname)
                    End If
                Next m_Site
            End If
        Next i
    End If
    
    ''''<MUST>
    If (FuseType = eFuse_ECID) Then
        ECIDFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_CFG) Then
        CFGFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_UID) Then
        UIDFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_SEN) Then
        SENFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_MON) Then
        MONFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_UDR) Then
        UDRFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_UDRE) Then
        UDRE_Fuse = m_Fuseblock
    ElseIf (FuseType = eFuse_UDRP) Then
        UDRP_Fuse = m_Fuseblock
    ElseIf (FuseType = eFuse_CMP) Then
        CMPFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_CMPE) Then
        CMPE_Fuse = m_Fuseblock
    ElseIf (FuseType = eFuse_CMPP) Then
        CMPP_Fuse = m_Fuseblock
    End If

    ''''debug purpose to check if the data are set to Read structure.
    If (False) Then
        For i = 0 To UBound(m_Fuseblock.Category)
            For Each m_Site In TheExec.sites
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.Decimal = " & m_Fuseblock.Category(i).Read.Decimal
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.Value   = " & m_Fuseblock.Category(i).Read.Value
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.HexStr  = " & m_Fuseblock.Category(i).Read.HexStr
            Next m_Site
        Next i
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'auto_eFuse_setReadData_forDatalog
''''201901XX Update
Public Function auto_eFuse_setReadData_forDatalog(ByVal FuseType As eFuseBlockType) As Boolean
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_setReadData_forDatalog"
   
    Dim m_Site As Variant
    Dim i As Long, j As Long
    Dim m_catename As String
    Dim m_startbit As Long
    Dim m_stopbit As Long
    Dim m_bitwidth As Long
    Dim m_stage As String
    Dim m_algorithm As String
''    Dim m_dec As New SiteVariant
''    Dim m_bitSum As New SiteLong
    Dim m_bitWave As New DSPWave

    Dim m_readDecArr() As Double
    Dim m_dblBitArr() As Long
    Dim m_bitArr() As Long
    Dim m_FuseRead As EFuseCategoryParamResultSyntax
    Dim m_Fuseblock As EFuseCategorySyntax
'    Dim m_bitstrM As String
'    Dim m_bitStrL As String
'    Dim m_hexStr As String
    Dim m_resolution As Double
    
''----New
    Dim m_dfreal As String
    Dim m_calc_mode As Long
    Dim mSD_dec As New SiteDouble
    Dim mSD_val As New SiteDouble
    Dim mSL_bitSum As New SiteLong
    Dim mSV_bitStrM As New SiteVariant
    Dim mSV_bitStrL As New SiteVariant
    Dim mSV_hexStr As New SiteVariant
    Dim mSB_cmpResult As New SiteBoolean
    Dim mSV_decimal As New SiteVariant
    Dim mSV_value As New SiteVariant
    
    Dim m_readDecCateWave As New DSPWave
    Dim mSD_readDecArr() As New SiteDouble
    Dim m_DoubleBitWave As New DSPWave
    Dim m_doubleBitArr() As Long
''---

    If (FuseType = eFuse_ECID) Then
        m_Fuseblock = ECIDFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_ECID_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_ECID_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_CFG) Then
        m_Fuseblock = CFGFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_CFG_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_CFG_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_UID) Then
          m_Fuseblock = UIDFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_UID_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_UID_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_SEN) Then
        m_Fuseblock = SENFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_SEN_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_SEN_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_MON) Then
        m_Fuseblock = MONFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_MON_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_MON_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_UDR) Then
        m_Fuseblock = UDRFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_UDR_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_UDR_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_UDRE) Then
        m_Fuseblock = UDRE_Fuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_UDRE_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_UDRE_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_UDRP) Then
        m_Fuseblock = UDRP_Fuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_UDRP_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_UDRP_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_CMP) Then
        m_Fuseblock = CMPFuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_CMP_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_CMP_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_CMPE) Then
        m_Fuseblock = CMPE_Fuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_CMPE_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_CMPE_Read_DoubleBitWave.Copy
        Next m_Site
    ElseIf (FuseType = eFuse_CMPP) Then
        m_Fuseblock = CMPP_Fuse
        For Each m_Site In TheExec.sites
            m_readDecCateWave = gDW_CMPP_Read_Decimal_Cate.Copy
            m_DoubleBitWave = gDW_CMPP_Read_DoubleBitWave.Copy
        Next m_Site
    End If

    ReDim mSD_readDecArr(UBound(m_Fuseblock.Category))
    
    For Each m_Site In TheExec.sites
        m_readDecArr = m_readDecCateWave.Data
        For i = 0 To UBound(m_readDecArr)
            mSD_readDecArr(i) = m_readDecArr(i)
        Next i
    Next m_Site
    
'    If (fusetype = eFuse_ECID) Then
'
'    Else
        For i = 0 To UBound(m_Fuseblock.Category)
            With m_Fuseblock.Category(i)
                m_catename = .Name
                m_stage = LCase(.Stage)
                If (.LSBbit > .MSBbit) Then
                    m_startbit = .MSBbit ''''<NOTICE>
                    m_stopbit = .LSBbit
                Else
                    m_startbit = .LSBbit ''''<NOTICE>
                    m_stopbit = .MSBbit
                End If
                m_bitwidth = .BitWidth
                m_algorithm = LCase(.algorithm)
                m_resolution = .Resoultion
                m_dfreal = LCase(.Default_Real)
            End With

            For Each m_Site In TheExec.sites
                m_bitWave = m_DoubleBitWave.Select(m_startbit, 1, m_bitwidth).Copy
                mSL_bitSum = m_bitWave.CalcSum
                m_Fuseblock.Category(i).Read.BitArrWave = m_bitWave.Copy
                Call auto_eFuse_bitWave_to_binStr_HexStr(m_bitWave, mSV_bitStrM, mSV_hexStr, True, False)
                mSV_bitStrL = StrReverse(mSV_bitStrM)
            Next m_Site

            With m_Fuseblock.Category(i).Read
                .BitStrM = mSV_bitStrM
                .BitStrL = mSV_bitStrL
                .BitSummation = mSL_bitSum
            End With
        Next i
'    End If
    
    ''''<MUST>
    If (FuseType = eFuse_ECID) Then
        ECIDFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_CFG) Then
        CFGFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_UID) Then
        UIDFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_SEN) Then
        SENFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_MON) Then
        MONFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_UDR) Then
        UDRFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_UDRE) Then
        UDRE_Fuse = m_Fuseblock
    ElseIf (FuseType = eFuse_UDRP) Then
        UDRP_Fuse = m_Fuseblock
    ElseIf (FuseType = eFuse_CMP) Then
        CMPFuse = m_Fuseblock
    ElseIf (FuseType = eFuse_CMPE) Then
        CMPE_Fuse = m_Fuseblock
    ElseIf (FuseType = eFuse_CMPP) Then
        CMPP_Fuse = m_Fuseblock
    End If

    ''''debug purpose to check if the data are set to Read structure.
    If (False) Then
        For i = 0 To UBound(m_Fuseblock.Category)
            For Each m_Site In TheExec.sites
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.Decimal = " & m_Fuseblock.Category(i).Read.Decimal
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.Value   = " & m_Fuseblock.Category(i).Read.Value
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.HexStr  = " & m_Fuseblock.Category(i).Read.HexStr
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.BitstrM = " & m_Fuseblock.Category(i).Read.BitStrM
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " Read.BitstrL = " & m_Fuseblock.Category(i).Read.BitStrL
                TheExec.Datalog.WriteComment "Site(" & m_Site & ") " & m_Fuseblock.Category(i).Name & " BitSummation = " & m_Fuseblock.Category(i).Read.BitSummation
                TheExec.Datalog.WriteComment "-----------------------"
            Next m_Site
        Next i
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''201808XX New, <NOTICE> still need to update later on
''''201812XX update method
'Public Function auto_eFuse_Simulate_fromWrite2CapWave(ByVal FuseType As eFuseBlockType, ByVal simBlank As SiteLong, ByRef outcapWave As DSPWave) As Boolean
Public Function auto_eFuse_Simulate_fromWrite2CapWave(ByVal FuseType As eFuseBlockType, ByVal simBlank As SiteLong, _
                                                      ByRef outcapWave As DSPWave, ByVal ReverseFlag As Boolean) As Boolean

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Simulate_fromWrite2CapWave"

    ''''------------------------------------------------------------------------------
    ''''  simBlank: simulate blank condition to decide which stage bit flag to be used
    ''''       = 0: means that all bits blank=True as early stage bits
    ''''       = 1: means that simulate those bits (stage <  job)
    ''''       = 2: means that simulate those bits (stage <= job)
    ''''------------------------------------------------------------------------------
    
    Dim i As Long
    Dim m_Site As Variant
    Dim m_stage As String
    Dim m_algorithm As String
    Dim m_defreal As String
    Dim m_crc_idx As Long
    Dim m_bitwidth As Long
    Dim m_crc_stage As String

    Dim m_pgmWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_tmpBitFlagWave As New DSPWave

    Dim m_Fuseblock As EFuseCategorySyntax
    
    Dim m_SampleSize As Long
    
    m_SampleSize = gDL_BitsPerBlock
    
    If (FuseType = eFuse_ECID) Then
        m_Fuseblock = ECIDFuse
    ElseIf (FuseType = eFuse_CFG) Then
        m_Fuseblock = CFGFuse
    ElseIf (FuseType = eFuse_UID) Then
        m_Fuseblock = UIDFuse
    ElseIf (FuseType = eFuse_SEN) Then
        m_Fuseblock = SENFuse
    ElseIf (FuseType = eFuse_MON) Then
        m_Fuseblock = MONFuse
    ElseIf (FuseType = eFuse_UDR) Then
        m_Fuseblock = UDRFuse
        m_SampleSize = gL_USI_DigSrcBits_Num
    ElseIf (FuseType = eFuse_UDRE) Then
        m_Fuseblock = UDRE_Fuse
        m_SampleSize = gL_UDRE_USI_DigSrcBits_Num
    ElseIf (FuseType = eFuse_UDRP) Then
        m_Fuseblock = UDRP_Fuse
        m_SampleSize = gL_UDRP_USI_DigSrcBits_Num
    ElseIf (FuseType = eFuse_CMP) Then
        m_Fuseblock = CMPFuse
        m_SampleSize = gL_CMP_DigCapBits_Num
    ElseIf (FuseType = eFuse_CMPE) Then
        m_Fuseblock = CMPE_Fuse
        m_SampleSize = gL_CMPE_DigCapBits_Num
    ElseIf (FuseType = eFuse_CMPP) Then
        m_Fuseblock = CMPP_Fuse
        m_SampleSize = gL_CMPE_DigCapBits_Num
    End If
    
     m_pgmWave.CreateConstant 0, m_SampleSize, DspLong
    
    For Each site In TheExec.sites
        gDW_Pgm_RawBitWave = m_pgmWave.Copy
    Next site

    If (FuseType = eFuse_ECID) Then
        Dim LotTmp As String
        Dim m_tmpwfid As String
        Dim m_len As Long
        Dim Loc_dash As Long:: Loc_dash = 1
        
        LotTmp = Trim(UCase(TheExec.Datalog.Setup.LotSetup.LotID))
        m_tmpwfid = Trim(CStr(TheExec.Datalog.Setup.WaferSetup.ID))
        m_len = Len(LotTmp)
        Loc_dash = InStr(1, LotTmp, "-")
'        If Loc_dash <> 0 Then
            'LotID = Mid(LotTmp, 1, Loc_dash - 1)
'        Else
            LotID = LotTmp
'        End If
        
        Call auto_eFuse_LotID_to_setWriteVariable(LotID)
        

        WaferID = CLng(m_tmpwfid)
        HramLotId = LotID
        HramWaferId = WaferID
        Call auto_eFuse_SetWriteVariable_SiteAware("ECID", "Wafer_ID", HramWaferId, False)
        
        For Each site In TheExec.sites
            XCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetXCoord(site)
            YCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetYCoord(site)
            Call auto_eFuse_SetWriteDecimal("ECID", "X_Coordinate", XCoord(site), False, False)
            Call auto_eFuse_SetWriteDecimal("ECID", "Y_Coordinate", YCoord(site), False, False)
        Next site
        
        
        ''''update it later
        For i = 0 To UBound(m_Fuseblock.Category) - 1
            With m_Fuseblock.Category(i)
                m_stage = LCase(.Stage)
                m_algorithm = LCase(.algorithm)
                m_defreal = LCase(.Default_Real)
                If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then ''''simulate bits Job>=Stage
                    If (m_algorithm = "crc") Then
                        ''''skip here and process in the last one
                        m_crc_idx = i
                        m_crc_stage = m_stage
                    ElseIf (m_defreal = "real" Or m_defreal = "bincut") Then
'                        If (m_algorithm = "vddbin") Then
'                            Call auto_eFuse_Vddbin_bincut_setWriteData(FuseType, i)
'                        End If
                        ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
                        Call rundsp.eFuse_updatePgmWave_byCategory(.BitIndexWave, .Write.BitArrWave)
                    End If
                End If
            End With

            ''''process CRC simulation (job>=crc_stage)
            If (i = UBound(m_Fuseblock.Category) And auto_eFuse_check_Job_cmpare_Stage(m_crc_stage) >= 0) Then
                With m_Fuseblock.Category(m_crc_idx)
                    ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
                    m_bitwidth = .BitWidth
                    Call rundsp.eFuse_updatePgmWave_CRCbits_Simulation(FuseType, m_bitwidth, .BitIndexWave)
                End With
            End If
        Next i

        Call rundsp.eFuse_Sim_Gen_32Bits_CapWave(FuseType, simBlank, outcapWave, True)

    Else ''''If (fusetype = eFuse_CFG) Then
        ''''except ECID Fuse, other blocks could be here
        For i = 0 To UBound(m_Fuseblock.Category)
            With m_Fuseblock.Category(i)
                m_stage = LCase(.Stage)
                m_algorithm = LCase(.algorithm)
                m_defreal = LCase(.Default_Real)
                If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0) Then ''''simulate bits Job>=Stage
                    If (m_algorithm = "crc") Then
                        ''''skip here and process in the last one
                        m_crc_idx = i
                        m_crc_stage = m_stage
                    ElseIf (m_defreal = "real" Or m_defreal = "bincut") Then
                        If (m_algorithm = "vddbin") Then
                            Call auto_eFuse_Vddbin_bincut_setWriteData(FuseType, i)
                        End If
                        ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
                        Call rundsp.eFuse_updatePgmWave_byCategory(.BitIndexWave, .Write.BitArrWave)
                    End If
                End If
            End With
            
            ''''process CRC simulation (job>=crc_stage)
            If (i = UBound(m_Fuseblock.Category) And auto_eFuse_check_Job_cmpare_Stage(m_crc_stage) >= 0) Then
            'If (i = UBound(m_Fuseblock.Category) - 1 And auto_eFuse_check_Job_cmpare_Stage(m_crc_stage) >= 0) Then
                With m_Fuseblock.Category(m_crc_idx)
                    ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
                    m_bitwidth = .BitWidth
                    ''''<MUST> using this call
                    Call rundsp.eFuse_updatePgmWave_CRCbits_Simulation(FuseType, m_bitwidth, .BitIndexWave)
                End With
            End If
        Next i
        
'    If (FuseType = eFuse_ECID) Then
'         ECIDFuse = m_Fuseblock
'    ElseIf (FuseType = eFuse_CFG) Then
'         CFGFuse = m_Fuseblock
'    ElseIf (FuseType = eFuse_UID) Then
'    ElseIf (FuseType = eFuse_SEN) Then
'    ElseIf (FuseType = eFuse_MON) Then
'    ElseIf (FuseType = eFuse_UDR) Then
'        UDRFuse = m_Fuseblock
'    End If
    End If

    'Call rundsp.eFuse_Sim_Gen_32Bits_CapWave(FuseType, simBlank, outcapWave)
    Call rundsp.eFuse_Sim_Gen_32Bits_CapWave(FuseType, simBlank, outcapWave, ReverseFlag)

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''201808XX New
Public Function auto_eFuse_print_capWave32Bits(fuseblock As eFuseBlockType, ByVal CapWave As DSPWave, Optional showPrint As Boolean = False) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_print_capWave32Bits"
    
    Dim i As Long, j As Long, k As Long
    Dim m_Site As Variant
    Dim m_size As Long
    Dim m_tmpStr As String
    Dim m_tmp32Str As String
    Dim m_len As Long
    Dim m_capArr() As Long ''''<NOTICE> type is double for capWave
    Dim m_sgBitArr() As Long
    Dim m_SingleBitWave As New DSPWave

    If (showPrint = False) Then Exit Function

    For Each m_Site In TheExec.sites
        m_size = CapWave.SampleSize
        If (m_size < 10000) Then m_len = 4
        If (m_size < 1000) Then m_len = 3
        If (m_size < 100) Then m_len = 2
        Exit For
    Next m_Site

    Select Case fuseblock
        Case eFuse_ECID:
            Call rundsp.eFuse_DspWave_Copy(gDW_ECID_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_CFG:
            Call rundsp.eFuse_DspWave_Copy(gDW_CFG_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_UID:
            Call rundsp.eFuse_DspWave_Copy(gDW_UID_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_UDR:
            Call rundsp.eFuse_DspWave_Copy(gDW_UDR_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_UDRE:
            Call rundsp.eFuse_DspWave_Copy(gDW_UDRE_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_UDRP:
            Call rundsp.eFuse_DspWave_Copy(gDW_UDRP_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_SEN:
            Call rundsp.eFuse_DspWave_Copy(gDW_SEN_Read_SingleBitWave, m_SingleBitWave)
        Case eFuse_MON:
            Call rundsp.eFuse_DspWave_Copy(gDW_MON_Read_SingleBitWave, m_SingleBitWave)
    End Select

'    TheExec.Datalog.WriteComment vbCrLf & ">>> Start <<< " + funcName
'    For Each m_Site In TheExec.Sites
'        m_capArr = capWave(m_Site).Data
'        m_sgBitArr = m_singleBitWave.Data
'        TheExec.Datalog.WriteComment "Site(" + CStr(m_Site) + ")"
'        m_tmpStr = ""
'        For i = 0 To m_size - 1
'            m_tmpStr = "CapWave.Element(" + FormatNumeric(i, m_len) + ") = " + FormatNumeric(m_capArr(i), -16)
'            m_tmp32Str = ""
'            For j = 0 To 31
'                k = i * 32 + j
'                m_tmp32Str = CStr(m_sgBitArr(k)) + m_tmp32Str
'                If (j = 15) Then m_tmp32Str = " " + m_tmp32Str
'            Next j
'            m_tmpStr = m_tmpStr + FormatNumeric(m_tmp32Str, 35)
'            TheExec.Datalog.WriteComment m_tmpStr
'        Next i
'        TheExec.Datalog.WriteComment ""
'    Next m_Site
'    TheExec.Datalog.WriteComment ">>> End <<< " & vbCrLf

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''201809XX update for the case with the different value by Site
Public Function eFuse_DSSC_SetupDigSrcWave(patt As String, DigSrcPin As PinList, SignalName As String, ByVal srcWave As DSPWave)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "eFuse_DSSC_SetupDigSrcWave"

    Dim site As Variant
    Dim m_segsize As Long
    Dim WaveDef As String

    TheHdw.Patterns(patt).Load
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName
    
    For Each site In TheExec.sites
        m_segsize = srcWave.SampleSize
        ''''20170920 <NOTICE> if multiple apply this function call/sequence to avoid the following SrcWave to overwrite the previous one
        WaveDef = "WaveDef_" + SignalName + "_" & site
        TheExec.WaveDefinitions.CreateWaveDefinition WaveDef, srcWave, True

        With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName)
            .WaveDefinitionName = WaveDef
            .SampleSize = m_segsize
            .Amplitude = 1
            ''.LoadSamples ''''could waste TT and break PTE
            .LoadSettings
        End With
    Next site
    TheHdw.Wait 0.0001
    
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
    TheHdw.Wait 0.0001

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''''201808XX new method with DspWave, used at MarginRead
Public Function auto_eFuse_compare_Read_PgmBitWave(fuseblock As eFuseBlockType, ByVal bitFlag_mode As Long, ByVal CapWave As DSPWave, _
                                                   ByRef FBCount As SiteLong, ByRef cmpResult As SiteLong, _
                                                   Optional blank_stage As Boolean = True, _
                                                   Optional allBlank As Boolean = True, _
                                                   Optional serialCap As Boolean = False, Optional PatBitOrder As String = "bit0_bitLast")

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_compare_Read_PgmBitWave"

'    Dim m_caseFlag As New SiteBoolean
'    If (PatBitOrder = "bit0_bitLast") Then
'        m_caseFlag = True
'    ElseIf (PatBitOrder = "bitLast_bit0") Then
'        m_caseFlag = False
'    End If
    Dim m_caseFlag As New SiteBoolean:: m_caseFlag = False
    If (PatBitOrder = "bit0_bitLast") Then
        m_caseFlag = False
    ElseIf (PatBitOrder = "bitLast_bit0") Then
        m_caseFlag = True
    End If


    'If (serialCap = False) Then
        Call rundsp.eFuse_compare_Read_PgmBitWave(fuseblock, bitFlag_mode, CapWave, FBCount, cmpResult, m_caseFlag)
        'Call RunDSP.eFuse_compare_Read_PgmBitWave(fuseblock, bitFlag_mode, capWave, FBCount, cmpResult)
    'Else
        ''''Need to implement later
        ''''serial capture

        'Call RunDSP.eFuse_compare_Read_PgmBitWave(fuseblock, bitFlag_mode, CapWave, FBCount, cmpResult, m_caseFlag)
        'Call RunDSP.eFuse_compare_Read_PgmBitWave(fuseblock, bitFlag_mode, capWave, FBCount, cmpResult)
    
        'Call RunDSP.eFuse_Wave1bit_to_SingleDoubleBitWave(fuseblock, bitFlag_mode, m_caseFlag, capWave, FBCount, blank_stage, allblank)
        'Call RunDSP.eFuse_Wave1bit_to_SingleDoubleBitWave(m_caseFlag, capWave, singleBitWave, doubleBitWave, FBCount, blank_stage)
    'End If
    TheHdw.Wait 0.001
    
Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

''''201811XX add to save TT and improve PTE for the case with all default bits
Public Function auto_eFuse_Gen_PgmBitSrcWave_onlyDeaultBits(ByVal FuseType As eFuseBlockType, _
                                                            ByRef outSrcWave As DSPWave, ByRef result As SiteLong) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Gen_PgmBitSrcWave_onlyDeaultBits"

    Dim m_Site As Variant
    Dim i As Long, j As Long, k As Long
    Dim k1 As Long, k2 As Long
    Dim m_tmpValue As Long
    Dim m_sgbits As Long
    
    Dim expandWidth As Long
    Dim BitsPerRow As Long
    Dim ReadCycles As Long
    Dim BitsPerCycle As Long
    Dim BitsPerBlock As Long
    
    Dim m_tmpWave1 As New DSPWave
    Dim m_pgmrawBitWave As New DSPWave
    Dim m_defaultBitWave As New DSPWave
    Dim m_effbitFlagWave As New DSPWave
    Dim m_singleWave As New DSPWave
    Dim m_doubleWave As New DSPWave

    Dim m_size As Long
    Dim m_dbbitArr() As Long
    Dim m_sgBitArr() As Long
    Dim m_outSrcBitArr() As Long
    
    BitsPerRow = gDL_BitsPerRow
    ReadCycles = gDL_ReadCycles
    BitsPerCycle = gDL_BitsPerCycle
    BitsPerBlock = gDL_BitsPerBlock
    expandWidth = gDL_DigSrcRepeatN

    ''''<MUST> because all Sites have the same data.
    ''''Just do once only
For Each m_Site In TheExec.sites
    If (FuseType = eFuse_CFG) Then
        ''m_defaultBitWave = gDW_CFG_allDefaultBitWave.Copy
        ''m_effbitFlagWave = gDW_CFG_Stage_Early_BitFlag.Copy
        
        ''''to save TT, 20190122
        m_defaultBitWave.Data = gL_CFG_allDefaultBits_arr
        m_effbitFlagWave.Data = gL_CFG_stage_early_bitFlag_arr
        
    ElseIf (FuseType = eFuse_UDR) Then
        m_defaultBitWave.Data = gL_UDR_allDefaultBits_arr
        m_effbitFlagWave.Data = gL_UDR_stage_early_bitFlag_arr
    ElseIf (FuseType = eFuse_UDRE) Then
        m_defaultBitWave.Data = gL_UDRE_allDefaultBits_arr
        m_effbitFlagWave.Data = gL_UDRE_stage_early_bitFlag_arr
    ElseIf (FuseType = eFuse_UDRP) Then
        m_defaultBitWave.Data = gL_UDRP_allDefaultBits_arr
        m_effbitFlagWave.Data = gL_UDRP_stage_early_bitFlag_arr
    End If
    
    m_pgmrawBitWave = gDW_Pgm_RawBitWave.Copy

    ''''combine m_pgmrawBitWave with "OR" m_defaultBitWave
    m_tmpWave1 = m_pgmrawBitWave.BitwiseOr(m_defaultBitWave).Copy
    
    ''''gen effective Wave with "AND" m_effbitFlagWave
    m_doubleWave = m_tmpWave1.bitwiseand(m_effbitFlagWave).Copy
    m_dbbitArr = m_doubleWave.Data
    
    m_sgbits = BitsPerCycle * ReadCycles
    m_singleWave.CreateConstant 0, m_sgbits, DspLong
    
    If (gDL_eFuse_Orientation = 0) Then ''''Up2Down
        m_singleWave = m_doubleWave.repeat(2).Copy

    ElseIf (gDL_eFuse_Orientation = 1) Then ''''SingleUp, eFuse_1_Bit
        ''''doubleBitWave is equal to singleBitWave
        m_singleWave = m_doubleWave.Copy

    ElseIf (gDL_eFuse_Orientation = 2) Then ''''Right2Left, eFuse_2_Bit
        ''''-------------------------------------------------------------------------------
        ''''New Method
        ''''-------------------------------------------------------------------------------
        m_sgBitArr = m_singleWave.Data

        k = 0 ''''must be here
        For i = 0 To ReadCycles - 1     'EX: EcidReadCycle - 1   ''0...15(ECID)
            For j = 0 To BitsPerRow - 1 'EX: EcidBitsPerRow - 1  ''0...15(ECID)
                ''''k1: Right block
                ''''k2:  Left block
                k1 = (i * BitsPerCycle) + j ''<Important> Must use BitsPerCycle here
                k2 = (i * BitsPerCycle) + BitsPerRow + j

                ''''save TT here by using dataArr
                m_tmpValue = m_dbbitArr(k)
                m_sgBitArr(k1) = m_tmpValue
                m_sgBitArr(k2) = m_tmpValue
                k = k + 1
            Next j
        Next i
        m_singleWave.Data = m_sgBitArr ''''save TT
        ''''-------------------------------------------------------------------------------

    ''''Below is reserved for the future, there is NO any definition at present.
    ElseIf (gDL_eFuse_Orientation = 3) Then
    ElseIf (gDL_eFuse_Orientation = 4) Then
    ElseIf (gDL_eFuse_Orientation = 5) Then
    End If

    ''''Expand the outWave to Source Wave
    m_size = m_sgbits * expandWidth
    outSrcWave.CreateConstant 0, m_size, DspLong
    ReDim m_outSrcBitArr(m_size - 1)
    
    If (gDL_eFuse_Orientation <> 2) Then m_sgBitArr = m_singleWave.Data
    
    k = 0
    For i = 0 To m_sgbits - 1
        For j = 0 To expandWidth - 1
            k = i * expandWidth + j
            ''outSrcWave.Element(k) = m_singleWave.Element(i) ''''waste TT
            m_outSrcBitArr(k) = m_sgBitArr(i) ''''to save TT
        Next j
    Next i
    outSrcWave.Data = m_outSrcBitArr ''''to save TT
    
    ''''<MUST> because all Sites have the same data.
    ''''Just do once only, save TT and improve PTE here
    Exit For
Next m_Site

    ''''update for the specific sites
    For Each m_Site In TheExec.sites
        ''m_singleWave.Data = m_sgBitArr
        ''m_doubleWave.Data = m_dbbitArr
        outSrcWave.Data = m_outSrcBitArr ''''<MUST>
    Next m_Site

    If (FuseType = eFuse_CFG) Then
        For Each m_Site In TheExec.sites
            gDW_CFG_Pgm_SingleBitWave.Data = m_sgBitArr
            gDW_CFG_Pgm_DoubleBitWave.Data = m_dbbitArr
        Next m_Site
    ElseIf (FuseType = eFuse_UDR) Then
        For Each m_Site In TheExec.sites
            gDW_UDR_Pgm_SingleBitWave.Data = m_sgBitArr
            gDW_UDR_Pgm_DoubleBitWave.Data = m_dbbitArr
        Next m_Site
    ElseIf (FuseType = eFuse_UDRE) Then
        For Each m_Site In TheExec.sites
            gDW_UDRE_Pgm_SingleBitWave.Data = m_sgBitArr
            gDW_UDRE_Pgm_DoubleBitWave.Data = m_dbbitArr
        Next m_Site
    ElseIf (FuseType = eFuse_UDRP) Then
        For Each m_Site In TheExec.sites
            gDW_UDRP_Pgm_SingleBitWave.Data = m_sgBitArr
            gDW_UDRP_Pgm_DoubleBitWave.Data = m_dbbitArr
        Next m_Site
    End If

    result = 1

Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

''''201811XX update
Public Function auto_eFuse_Global_DSPWave_Reset() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Global_DSPWave_Reset"

    ''''---------------------------------------------
    '''' Clear the global DSP DSPWave
    ''''---------------------------------------------
    gDW_Pgm_RawBitWave.Clear
    gDW_ECID_CRC_calcBitsWave.Clear
    gDW_CFG_CRC_calcBitsWave.Clear
    gDW_UID_CRC_calcBitsWave.Clear
    gDW_SEN_CRC_calcBitsWave.Clear
    gDW_MON_CRC_calcBitsWave.Clear
    gDW_Pgm_BitWaveForCRCCalc.Clear
    gDW_Read_BitWaveForCRCCalc.Clear
    gDW_UID_CRC_calcBitsWave.Clear
    gDW_MON_CRC_calcBitsWave.Clear
    gDW_SEN_CRC_calcBitsWave.Clear
    ''''---------------------------------------------
    gDW_ECID_MSBBit_Cate.Clear
    gDW_ECID_LSBBit_Cate.Clear
    gDW_ECID_BitWidth_Cate.Clear
    gDW_ECID_DefaultReal_Cate.Clear
    gDW_ECID_Stage_BitFlag.Clear
    gDW_ECID_Stage_Early_BitFlag.Clear
    gDW_ECID_allDefaultBitWave.Clear
    gDW_ECID_Read_Decimal_Cate.Clear
    gDW_ECID_Read_cmpsgWavePerCyc.Clear
    gDW_ECID_Pgm_SingleBitWave.Clear
    gDW_ECID_Pgm_DoubleBitWave.Clear
    gDW_ECID_Read_SingleBitWave.Clear
    gDW_ECID_Read_DoubleBitWave.Clear
    gDW_ECID_StageLEQJob_BitFlag.Clear
    ''''---------------------------------------------
    gDW_CFG_MSBBit_Cate.Clear
    gDW_CFG_LSBBit_Cate.Clear
    gDW_CFG_BitWidth_Cate.Clear
    gDW_CFG_DefaultReal_Cate.Clear
    gDW_CFG_Stage_BitFlag.Clear
    gDW_CFG_Stage_Early_BitFlag.Clear
    gDW_CFG_Stage_Real_BitFlag.Clear
    gDW_CFG_allDefaultBitWave.Clear
    gDW_CFG_Read_Decimal_Cate.Clear
    gDW_CFG_Read_cmpsgWavePerCyc.Clear
    gDW_CFG_Pgm_SingleBitWave.Clear
    gDW_CFG_Pgm_DoubleBitWave.Clear
    gDW_CFG_Read_SingleBitWave.Clear
    gDW_CFG_Read_DoubleBitWave.Clear
    gDW_CFG_StageLEQJob_BitFlag.Clear
    gDW_CFG_SegFlag.Clear
    ''''---------------------------------------------
    gDW_MON_CRC_calcBitsWave.Clear
    gDW_SEN_CRC_calcBitsWave.Clear
    gDW_UID_CRC_calcBitsWave.Clear
Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function


''''201812XX update
Public Function auto_eFuse_Vddbin_bincut_setWriteData(ByVal FuseType As eFuseBlockType, ByVal index As Long) As Long
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Vddbin_bincut_setWriteData"

    Dim i As Long
    Dim m_Site As Variant
    Dim m_catename As String
    Dim m_catenameVbin As String
    Dim m_resolution As Double
    Dim mSD_vbinResult As New SiteDouble
    Dim mSV_decimal As New SiteVariant
    Dim m_tmpStr As String
    ''''---------------------------------------------
    Dim m_FuseStr As String
    Dim m_FuseWrite As EFuseCategoryParamResultSyntax
    Dim m_Fuseblock As EFuseCategorySyntax
    ''''---------------------------------------------

    i = index

    If (FuseType = eFuse_CFG) Then
        m_FuseStr = "CFG"
        m_Fuseblock = CFGFuse
        m_FuseWrite = CFGFuse.Category(i).Write
    ElseIf (FuseType = eFuse_UDR) Then
        m_FuseStr = "UDR"
        m_Fuseblock = UDRFuse
        m_FuseWrite = UDRFuse.Category(i).Write
    ElseIf (FuseType = eFuse_UDRE) Then
        m_FuseStr = "UDRE"
        m_Fuseblock = UDRE_Fuse
        m_FuseWrite = UDRE_Fuse.Category(i).Write
    ElseIf (FuseType = eFuse_UDRP) Then
        m_FuseStr = "UDRP"
        m_Fuseblock = UDRP_Fuse
        m_FuseWrite = UDRP_Fuse.Category(i).Write
    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (CFG,UDR,UDRE,UDRP)"
        GoTo errHandler
    End If

    m_catename = m_Fuseblock.Category(i).Name
    m_resolution = m_Fuseblock.Category(i).Resoultion
    
    If (m_resolution <= 0) Then
        m_tmpStr = funcName + ":: Please have positive resolution " & m_resolution & " (" + m_FuseStr + "::" + m_catename + ")"
        TheExec.AddOutput m_tmpStr
        TheExec.Datalog.WriteComment m_tmpStr
        GoTo errHandler
    End If
    ''''<Notice>
    ''''Here m_catename MUST be same as the content of Enum EcidVddBinningFlow
    ''''Ex: VDD_SRAM_P1 in (Enum EcidVddBinningFlow)
    'm_catenameVBin = "MS001" ''''<NOTICE> M8 uses MS001 on both power VDD_SOC and VDD_SOC_AON
    m_catenameVbin = m_catename

    mSD_vbinResult = VBIN_RESULT(VddBinStr2Enum(m_catenameVbin)).GRADEVDD
    
    If (TheExec.TesterMode = testModeOffline) Then
        For Each m_Site In TheExec.sites
            If (mSD_vbinResult = 0 Or mSD_vbinResult = -1) Then
                mSD_vbinResult = gD_BaseVoltage + m_resolution * auto_eFuse_GetWriteDecimal(m_FuseStr, m_catename, False)
            End If
        Next m_Site
    End If
    
    mSV_decimal = mSD_vbinResult.Subtract(gD_BaseVoltage).Divide(m_resolution)
    
    ''''<NOTICE and IMPORTANCE>
    For Each m_Site In TheExec.sites
        mSV_decimal = CeilingValue(mSV_decimal, 1)
        If mSD_vbinResult < 0 Then
            mSV_decimal = 0
            TheExec.Datalog.WriteComment "<WARNING> " + funcName + "::" + m_catename + _
                                         " Vbin_Result <0 (" & mSD_vbinResult & ") and force it to zero."
        End If
        If (mSV_decimal(m_Site) < 0) Then mSV_decimal(m_Site) = 0
    Next m_Site
    
    Call auto_eFuse_SetWriteVariable_SiteAware(m_FuseStr, m_catename, mSV_decimal)
'    With m_FuseWrite
'        .Value = mSD_vbinResult
'        .ValStr = .Value
'    End With

    If (FuseType = eFuse_CFG) Then
        'CFGFuse.Category(i).Write = m_FuseWrite
        CFGFuse.Category(i).Write.Value = mSD_vbinResult
        CFGFuse.Category(i).Write.ValStr = CFGFuse.Category(i).Write.Value
    ElseIf (FuseType = eFuse_UDR) Then
        'UDRFuse.Category(i).Write = m_FuseWrite
        UDRFuse.Category(i).Write.Value = mSD_vbinResult
        UDRFuse.Category(i).Write.ValStr = CFGFuse.Category(i).Write.Value
    ElseIf (FuseType = eFuse_UDRE) Then
        'UDRE_Fuse.Category(i).Write = m_FuseWrite
        UDRE_Fuse.Category(i).Write.Value = mSD_vbinResult
        UDRE_Fuse.Category(i).Write.ValStr = CFGFuse.Category(i).Write.Value
    ElseIf (FuseType = eFuse_UDRP) Then
        'UDRP_Fuse.Category(i).Write = m_FuseWrite
        UDRP_Fuse.Category(i).Write.Value = mSD_vbinResult
        UDRP_Fuse.Category(i).Write.ValStr = CFGFuse.Category(i).Write.Value
    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (CFG,UDR,UDRE,UDRP)"
        GoTo errHandler
    End If


Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

''''201812XX update
Public Function auto_eFuse_bitArr_to_binStr_HexStr(bitArr() As Long, BitStrM As String, HexStr As String, Optional mBit0IsLSB = True) As Long
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_bitArr_to_binStr_HexStr"
    
    Dim i As Long
    Dim m_bitStrM As String
    Dim m_HexStr As String
    
    m_bitStrM = ""
    m_HexStr = ""
    If (mBit0IsLSB) Then
        For i = 0 To UBound(bitArr)
            m_bitStrM = CStr(bitArr(i)) + m_bitStrM
        Next i
    Else
        ''''bit[0] is MSB
        For i = 0 To UBound(bitArr)
            m_bitStrM = m_bitStrM + CStr(bitArr(i))
        Next i
    End If
    
    m_HexStr = "0x" + auto_BinStr2HexStr(m_bitStrM, 1)
    
    ''''return the result
    BitStrM = m_bitStrM
    HexStr = m_HexStr

Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_eFuse_bitWave_to_binStr_HexStr(ByVal bitWave As DSPWave, BitStrM As SiteVariant, HexStr As SiteVariant, Optional mBit0IsLSB = True, Optional showPrint As Boolean = False) As Long
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_bitWave_to_binStr_HexStr"
    
    Dim m_Site As Variant
    Dim m_size As Long
    Dim m_bitArr() As Long
    
    Dim m_bitStrM As String
    Dim m_HexStr As String
    
    For Each m_Site In TheExec.sites
        m_bitArr = bitWave.Data
        m_size = bitWave.SampleSize
        Call auto_eFuse_bitArr_to_binStr_HexStr(m_bitArr, m_bitStrM, m_HexStr, mBit0IsLSB)
        BitStrM = m_bitStrM
        HexStr = m_HexStr
        If (showPrint) Then
            TheExec.Datalog.WriteComment "Site(" & m_Site & ")," & m_size & "bits [MSB...LSB] = [" + m_bitStrM + "] " + m_HexStr
        End If
    Next m_Site
    ''If (showPrint) Then TheExec.Datalog.WriteComment ""
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

'''''''''''-------------------------------------------------------------------
''''Just for Reference in Function auto_eFuse_DSSC_ReadDigCap_32bits_NEW()
'''''''''''-------------------------------------------------------------------
'''''''''''                                 3322222222221111 1111100000000000
''''''''''' Site (2)                        1098765432109876 5432109876543210
'''''''''''-------------------------------------------------------------------
''''CapWave.Element( 0) = 263458740         0000111110110100 0000111110110100
''''CapWave.Element( 1) = 537927696         0010000000010000 0010000000010000
''''CapWave.Element( 2) = 1343639574        0101000000010110 0101000000010110
''''CapWave.Element( 3) = 524296            0000000000001000 0000000000001000
''''CapWave.Element( 4) = 1137460172        0100001111001100 0100001111001100
''''CapWave.Element( 5) = 612967561         0010010010001001 0010010010001001
''''CapWave.Element( 6) = -182848231        1111010100011001 1111010100011001
''''CapWave.Element( 7) = -2147450880       1000000000000000 1000000000000000
''''CapWave.Element( 8) = 1062813529        0011111101011001 0011111101011001
''''CapWave.Element( 9) = 755838221         0010110100001101 0010110100001101
''''CapWave.Element(10) = 1957393579        0111010010101011 0111010010101011
''''CapWave.Element(11) = -1296911694       1011001010110010 1011001010110010
''''CapWave.Element(12) = -1157580032       1011101100000000 1011101100000000
''''CapWave.Element(13) = -859452219        1100110011000101 1100110011000101
''''CapWave.Element(14) = -187894580        1111010011001100 1111010011001100
''''CapWave.Element(15) = 0                 0000000000000000 0000000000000000
''''
'''''''''''-------------------------------------------------------------------
'''''''''''                                 3322222222221111 1111100000000000
''''''''''' Site (3)                        1098765432109876 5432109876543210
'''''''''''-------------------------------------------------------------------
''''CapWave.Element( 0) = 263458740         0000111110110100 0000111110110100
''''CapWave.Element( 1) = 537927696         0010000000010000 0010000000010000
''''CapWave.Element( 2) = 1880518678        0111000000010110 0111000000010110
''''CapWave.Element( 3) = 786444            0000000000001100 0000000000001100
''''CapWave.Element( 4) = -1645961756       1001110111100100 1001110111100100
''''CapWave.Element( 5) = 195431334         0000101110100110 0000101110100110
''''CapWave.Element( 6) = -220532006        1111001011011010 1111001011011010
''''CapWave.Element( 7) = -2147450880       1000000000000000 1000000000000000
''''CapWave.Element( 8) = 649668281         0010011010111001 0010011010111001
''''CapWave.Element( 9) = 745352301         0010110001101101 0010110001101101
''''CapWave.Element(10) = 1824222395        0110110010111011 0110110010111011
''''CapWave.Element(11) = -774712878        1101000111010010 1101000111010010
''''CapWave.Element(12) = -417667302        1110011100011010 1110011100011010
''''CapWave.Element(13) = -1717593697       1001100110011111 1001100110011111
''''CapWave.Element(14) = 1771661721        0110100110011001 0110100110011001
''''CapWave.Element(15) = 0                 0000000000000000 0000000000000000
'''''''''''-------------------------------------------------------------------

''''----------------------------------------------------------------------------------------
''''    Dim m_simWave As New DSPWave
''''    Dim m_cnt As Long
''''    m_cnt = 0
''''    m_simWave.CreateConstant 0, cycleNum, DspLong
''''    For Each Site In TheExec.Sites
''''        If (m_cnt = 0) Then
''''            ''''-------------------------------------------------------------------
''''            ''''                                 3322222222221111 1111100000000000
''''            ''''                                 1098765432109876 5432109876543210
''''            ''''-------------------------------------------------------------------
''''            m_simWave.Element(0) = -476716139  ''1110001110010101 1110001110010101
''''            m_simWave.Element(1) = -2145812455 ''1000000000011001 1000000000011001
''''            m_simWave.Element(2) = -1815440438 ''1001001111001010 1001001111001010
''''            m_simWave.Element(3) = 252645135   ''0000111100001111 0000111100001111
''''            m_simWave.Element(4) = -1717790308 ''1001100110011100 1001100110011100
''''            m_simWave.Element(5) = 251858691   ''0000111100000011 0000111100000011
''''            m_simWave.Element(6) = 1010580540  ''0011110000111100 0011110000111100
''''        Else
''''            ''''-------------------------------------------------------------------
''''            ''''                                 3322222222221111 1111100000000000
''''            ''''                                 1098765432109876 5432109876543210
''''            ''''-------------------------------------------------------------------
''''            m_simWave.Element(0) = -2145812455 ''1000000000011001 1000000000011001
''''            m_simWave.Element(1) = -1815440438 ''1001001111001010 1001001111001010
''''            m_simWave.Element(2) = 252645135   ''0000111100001111 0000111100001111
''''            m_simWave.Element(3) = -1717790308 ''1001100110011100 1001100110011100
''''            m_simWave.Element(4) = 251858691   ''0000111100000011 0000111100000011
''''            m_simWave.Element(5) = 1010580540  ''0011110000111100 0011110000111100
''''            m_simWave.Element(6) = -476716139  ''1110001110010101 1110001110010101
''''        End If
''''        m_cnt = m_cnt + 1
''''    Next Site
''''----------------------------------------------------------------------------------------

''''------------------------------------------------------------------------------------
''''Below, Just a trial but not using anymore
''''It could take a long time than host PC.
''''------------------------------------------------------------------------------------
''    If (fusetype = eFuse_ECID) Then
''        For i = 0 To UBound(ECIDFuse.Category) - 1
''            With ECIDFuse.Category(i)
''                m_catename = .Name
''                m_stage = LCase(.Stage)
''                m_startbit = .MSBbit ''''<NOTICE>
''                m_stopbit = .LSBbit
''            End With
''            Call RunDSP.eFuse_Get_ValueFromWaveXX(i, readDecCateWave, m_startbit, m_stopbit, doubleBitWave, m_dec, m_bitSum, m_bitWave)
''            TheHdw.Wait 10 * us ''''??? needed
''            ''''Here is just to get Read.Decimal/.BitSummation/.bitArrWave, to save the TTR
''            ''''for other properties of Read, they are leaved in the print section if needed.
''            ECIDFuse.Category(i).Read.Decimal = m_dec
''            ECIDFuse.Category(i).Read.BitSummation = m_bitSum
''            ECIDFuse.Category(i).Read.bitArrWave = m_bitWave
''        Next i
''    ElseIf (fusetype = eFuse_CFG) Then
''        For i = 0 To UBound(CFGFuse.Category)
''            With CFGFuse.Category(i)
''                m_catename = .Name
''                m_stage = LCase(.Stage)
''                m_startbit = .LSBbit ''''<NOTICE>
''                m_stopbit = .MSBbit
''            End With
''            Call RunDSP.eFuse_Get_ValueFromWaveXX(i, readDecCateWave, m_startbit, m_stopbit, doubleBitWave, m_dec, m_bitSum, m_bitWave)
''            TheHdw.Wait 10 * us ''''??? needed
''            ''''Here is just to get Read.Decimal/.BitSummation/.bitArrWave, to save the TTR
''            ''''for other properties of Read, they are leaved in the print section if needed.
''            CFGFuse.Category(i).Read.Decimal = m_dec
''            CFGFuse.Category(i).Read.BitSummation = m_bitSum
''            CFGFuse.Category(i).Read.bitArrWave = m_bitWave
''        Next i
''    End If
''''------------------------------------------------------------------------------------
Public Function auto_PrintBitMap(HramArray() As Long, _
                                 ByVal TotalCycleNumber As Long, _
                                 ByVal TotalBitNum As Long, _
                                 ByVal BitNumPerRow As Long, _
                                 ByVal FuseType As eFuseBlockType)
  
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_PrintBitMap"
    
    Dim Row_Str As String
    Dim L As Long, j As Long, k As Long, i As Long
    Dim max_expand As Long
    Dim K_divided As Long
    Dim Block As Long
    Dim headerStr As String
    Dim ss As Variant
    Dim m_PerRowSize As Long:: m_PerRowSize = 32
    
    ss = TheExec.sites.SiteNumber
    
    If (gB_eFuse_newMethod) Then
        Dim cmpRes_Arr() As Long
        Select Case FuseType
            Case eFuse_ECID:
                cmpRes_Arr = gDW_ECID_Read_cmpsgWavePerCyc(ss).Data
            Case eFuse_CFG:
                cmpRes_Arr = gDW_CFG_Read_cmpsgWavePerCyc(ss).Data
            Case eFuse_UID:
                cmpRes_Arr = gDW_UID_Read_cmpsgWavePerCyc(ss).Data
            Case eFuse_MON:
                cmpRes_Arr = gDW_MON_Read_cmpsgWavePerCyc(ss).Data
            Case eFuse_SEN:
                cmpRes_Arr = gDW_SEN_Read_cmpsgWavePerCyc(ss).Data
            Case eFuse_UDR:
                If (gDL_eFuse_Orientation = 1 And BitNumPerRow = 1) Then _
                    TotalCycleNumber = TotalCycleNumber / m_PerRowSize
                'cmpRes_Arr = gDW_UDR_Read_cmpsgWavePerCyc(ss).Data
            Case eFuse_UDRE:
                If (gDL_eFuse_Orientation = 1 And BitNumPerRow = 1) Then _
                    TotalCycleNumber = TotalCycleNumber / m_PerRowSize
                'cmpRes_Arr = gDW_UDRE_Read_cmpsgWavePerCyc(ss).Data
            Case eFuse_UDRP
                If (gDL_eFuse_Orientation = 1 And BitNumPerRow = 1) Then _
                    TotalCycleNumber = TotalCycleNumber / m_PerRowSize
                'cmpRes_Arr = gDW_UDRP_Read_cmpsgWavePerCyc(ss).Data
            Case eFuse_CMP:
                If (gDL_eFuse_Orientation = 1 And BitNumPerRow = 1) Then _
                    TotalCycleNumber = TotalCycleNumber / m_PerRowSize
               ' cmpRes_Arr = gDW_CMP_Read_cmpsgWavePerCyc(ss).Data
            Case eFuse_CMPE:
                If (gDL_eFuse_Orientation = 1 And BitNumPerRow = 1) Then _
                    TotalCycleNumber = TotalCycleNumber / m_PerRowSize
                'cmpRes_Arr = gDW_CMPE_Read_cmpsgWavePerCyc(ss).Data
            Case eFuse_CMPP:
                If (gDL_eFuse_Orientation = 1 And BitNumPerRow = 1) Then _
                    TotalCycleNumber = TotalCycleNumber / m_PerRowSize
                'cmpRes_Arr = gDW_CMPP_Read_cmpsgWavePerCyc(ss).Data
            Case Else:
                GoTo errHandler
        End Select
    End If
    
    TheExec.Datalog.WriteComment ""
    
    Row_Str = ""
    
    headerStr = "====== Efuse Data read from DSSC (Chip internal ORed data) "
    headerStr = headerStr + "( " + TheExec.DataManager.instanceName + " )" + " (Site" + CStr(ss) + ")"
    headerStr = headerStr + "============"

    If (BitNumPerRow = 8) Then
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "           00000000"
        TheExec.Datalog.WriteComment "           76543210"
        TheExec.Datalog.WriteComment ""
    ElseIf (BitNumPerRow = 16) Then
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "           1111110000000000"
        TheExec.Datalog.WriteComment "           5432109876543210"
        TheExec.Datalog.WriteComment ""
    ElseIf (BitNumPerRow = 32 And (gS_EFuse_Orientation = "RIGHT2LEFT")) Then
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "           3322222222221111 1111110000000000  C"
        TheExec.Datalog.WriteComment "           1098765432109876 5432109876543210  M"
        TheExec.Datalog.WriteComment "           ---------------- ----------------  P"
    
    ElseIf (BitNumPerRow = 32) Then
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "           33222222222211111111110000000000"
        TheExec.Datalog.WriteComment "           10987654321098765432109876543210"
        TheExec.Datalog.WriteComment ""
    Else
        TheExec.Datalog.WriteComment headerStr
        TheExec.Datalog.WriteComment "Wrong BitNumPerRow(" + CStr(BitNumPerRow) + "), only 8, 16 and 32 support"
        TheExec.Datalog.WriteComment "           33222222222211111111110000000000"
        TheExec.Datalog.WriteComment "           10987654321098765432109876543210"
        TheExec.Datalog.WriteComment ""
    End If

    k = 0
    If (BitNumPerRow = 32 And (gS_EFuse_Orientation = "RIGHT2LEFT")) Then
        For i = 0 To TotalCycleNumber - 1
            For j = 0 To BitNumPerRow - 1
                If (j = 16) Then
                    ''''ading one space " " between bit16 and bit15
                    Row_Str = CStr(HramArray(k)) + " " + Row_Str
                Else
                    Row_Str = CStr(HramArray(k)) + Row_Str
                End If
                k = k + 1
            Next j
        Next i
        ''''because adding one speace " " per row, so total bit number should add cycle numbers tp parse Row_Str
        BitNumPerRow = BitNumPerRow + 1
        TotalBitNum = TotalBitNum + TotalCycleNumber
    Else
        If (FuseType = eFuse_ECID Or FuseType = eFuse_CFG Or FuseType = eFuse_MON Or FuseType = eFuse_SEN) Then m_PerRowSize = BitNumPerRow
        If (TotalBitNum Mod m_PerRowSize <> 0) Then TotalCycleNumber = TotalCycleNumber + 1
        For i = 0 To TotalCycleNumber - 1
            For j = 0 To m_PerRowSize - 1
                If (k < TotalBitNum) Then
                    Row_Str = CStr(HramArray(k)) + Row_Str
                Else
                    Row_Str = "0" + Row_Str
                End If
                k = k + 1
            Next j
        Next i
    End If


    If (FuseType <> eFuse_ECID And FuseType <> eFuse_CFG And FuseType <> eFuse_MON And FuseType <> eFuse_SEN And FuseType <> eFuse_UID) Then
        For i = 1 To TotalCycleNumber
            TheExec.Datalog.WriteComment "Row = " & Format((i - 1), "000") & ": " & Mid(Row_Str, (m_PerRowSize * TotalCycleNumber - (m_PerRowSize * i)) + 1, m_PerRowSize)
        Next i
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT" And gB_eFuse_newMethod = True) Then
        For i = 1 To TotalCycleNumber
            TheExec.Datalog.WriteComment "Row = " & Format((i - 1), "000") & ": " & Mid(Row_Str, (TotalBitNum - (BitNumPerRow * i)) + 1, BitNumPerRow) & "  " & cmpRes_Arr(i - 1)
        Next i
    End If

    TheExec.Datalog.WriteComment "====== End of Efuse Data read from DSSC ============"
    TheExec.Datalog.WriteComment ""

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'auto_eFuse_SyntaxCheck(m_Fusetype, condStr)
Public Function auto_eFuse_SyntaxCheck(FuseType As eFuseBlockType, Optional CondtionStr As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_SyntaxCheck"
    
    Dim i As Long
    Dim m_tsname As String
    Dim m_tsName0 As String
    Dim m_stage As String
    Dim m_catename As String
    Dim m_catenameVbin As String
    Dim m_bitStrM As String
    Dim m_bitwidth As Long
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_decimal As Variant ''20160506 update, was Long
    Dim m_decimal_ids As Double
    Dim m_bitsum As Long
    Dim m_value As Variant
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim TmpVal As Variant
    Dim vbinIdx As Long
    Dim vbinflag As Long
    Dim tmpStr As String
    Dim mem_Config_code As String
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim m_Pmode As Long
    Dim MaxLevelIndex As Long
    Dim m_testValue As Variant
    
    ''''20170630 add
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_HexStr As String
    Dim m_bitStrM_Tab As String
    Dim m_hexStr_Tab As String
    Dim m_bitStr32M_Read As String
    
    ''''201811XX
    Dim m_customUnit As String
    Dim mSL_bitSum As New SiteLong
    Dim mSV_value As New SiteVariant
    Dim mSV_decimal As New SiteVariant
    Dim mSV_bitStrM As New SiteVariant
    Dim mSV_hexStr As New SiteVariant
    Dim mSL_valueSum As New SiteLong
    Dim m_findTMPS_flag As Boolean
    Dim m_unitType As UnitType
    Dim m_scale As tlScaleType
    Dim m_crchexStr As String
    Dim m_Fuseblock As EFuseCategorySyntax
    Dim m_FuseBankStr As String
    Dim mS_Read_calcCRC_hexStr As New SiteVariant
    
    Dim m_GetHLlimit As Boolean:: m_GetHLlimit = False
    Dim mSV_lolmt As New SiteVariant
    Dim mSV_hilmt As New SiteVariant
    
    Dim m_CalcSum As Long
    
    If (FuseType = eFuse_ECID) Then
        m_Fuseblock = ECIDFuse
        m_FuseBankStr = "ECID"
        mS_Read_calcCRC_hexStr = gS_ECID_Read_calcCRC_hexStr
    ElseIf (FuseType = eFuse_CFG) Then
        m_Fuseblock = CFGFuse
        m_FuseBankStr = "CFG"
        mS_Read_calcCRC_hexStr = gS_CFG_Read_calcCRC_hexStr
    ElseIf (FuseType = eFuse_UID) Then
    ElseIf (FuseType = eFuse_SEN) Then
        m_Fuseblock = SENFuse
        m_FuseBankStr = "SEN"
        mS_Read_calcCRC_hexStr = gS_SEN_Read_calcCRC_hexStr
    ElseIf (FuseType = eFuse_MON) Then
        m_Fuseblock = MONFuse
        m_FuseBankStr = "MON"
        mS_Read_calcCRC_hexStr = gS_MON_Read_calcCRC_hexStr
    ElseIf (FuseType = eFuse_UDR) Then
        m_Fuseblock = UDRFuse
        m_FuseBankStr = "UDR"
    ElseIf (FuseType = eFuse_UDRE) Then
        m_Fuseblock = UDRE_Fuse
        m_FuseBankStr = "UDRE"
    ElseIf (FuseType = eFuse_UDRP) Then
        m_Fuseblock = UDRP_Fuse
        m_FuseBankStr = "UDRP"
    ElseIf (FuseType = eFuse_CMP) Then
        m_Fuseblock = CMPFuse
        m_FuseBankStr = "CMP"
        m_GetHLlimit = True
    ElseIf (FuseType = eFuse_CMPE) Then
        m_Fuseblock = CMPE_Fuse
        m_FuseBankStr = "CMPE"
        m_GetHLlimit = True
    ElseIf (FuseType = eFuse_CMPP) Then
        m_Fuseblock = CMPP_Fuse
        m_FuseBankStr = "CMPP"
        m_GetHLlimit = True
    End If
    
    For i = 0 To UBound(m_Fuseblock.Category)
        With m_Fuseblock.Category(i)
            m_stage = LCase(.Stage)
            m_catename = .Name
            'Debug.Print m_catename
            m_algorithm = LCase(.algorithm)
            'm_value = .Read.Value(Site)
            m_lolmt = .LoLMT
            m_hilmt = .HiLMT
            m_bitwidth = .BitWidth
            m_defval = .DefaultValue
            m_resolution = .Resoultion
            m_defreal = LCase(.Default_Real)
            If (m_GetHLlimit = True) Then
                mSV_hilmt = .Read.Decimal
                mSV_lolmt = .Read.Decimal
            End If
            ''''-----------------------------------
            mSL_bitSum = .Read.BitSummation
            mSV_decimal = .Read.Decimal
            mSV_bitStrM = .Read.BitStrM
            mSV_hexStr = .Read.HexStr
            mSV_value = .Read.Value
            ''''-----------------------------------
        End With
        
        m_tsname = Replace(m_catename, " ", "_") ''''20151028, benefit for the script
        m_unitType = unitNone
        m_customUnit = ""
        m_scale = scaleNoScaling ''=scaleNone ''''default

        ''''Here it presents the CFG_Cond_[Flag Selected] syntax check
        If (i = 0 And m_FuseBankStr = "CFG") Then
            m_tsName0 = "CFG_Cond_" + gS_cfgFlagname '''' + "_Read_" + gS_CFG_Cond_Read_pkgname(Site)) then need siteloop for the tsname per site
            mSV_value = gL_CFG_Cond_compResult
            TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=0, hiVal:=0, Tname:=m_tsName0
        End If
        
        ''If (m_stage = gS_JobName) Then ''by Job
        'If (True) Then
        If (((m_stage = gS_JobName) And CondtionStr = "cp1_early") Or CondtionStr <> "cp1_early") Then ''by Job
                
            ''''<NOTICE> 20160108
            If (m_GetHLlimit <> True) Then
                Call auto_eFuse_chkLoLimit(m_FuseBankStr, i, m_stage, m_lolmt)
                Call auto_eFuse_chkHiLimit(m_FuseBankStr, i, m_stage, m_hilmt)
            End If
            
'            If (FuseType = eFuse_CMP Or FuseType = eFuse_CMPE Or FuseType = eFuse_CMPP) Then
'                m_lolmt = m_Fuseblock.Category(i).LoLMT
'                m_hilmt = m_Fuseblock.Category(i).HiLMT
'            End If
            
            ''''<TRY>dummy Test
            If (m_algorithm = "ids") Then
                m_unitType = unitCustom
                m_customUnit = "mA"
                If (CDbl(m_lolmt) = 0# And CDbl(m_hilmt) = 0#) Then ''''Need to check
                    m_lolmt = 1# * m_resolution
                    m_hilmt = ((2 ^ m_bitwidth) - 1) * m_resolution
                Else
                    If (CDbl(m_lolmt) = 0#) Then m_lolmt = 1# * m_resolution  '0 means nothing, can not be acceptable
                End If
            
            ElseIf (m_algorithm = "tmps") Then
                ''''--------------------------------------------------------------------------------------------
                ''''20160503 update for the 2nd way to prevent for the temp sensor trim will all zero values
                ''''201812XX update
                m_findTMPS_flag = True
                mSL_valueSum = mSL_valueSum.Add(mSV_decimal)
                ''''--------------------------------------------------------------------------------------------

            ElseIf (m_algorithm = "base" And m_defreal Like "safe*voltage") Then
                m_unitType = unitCustom
                m_customUnit = "mV"

            ElseIf (m_algorithm = "vddbin") Then
                m_unitType = unitCustom
                m_customUnit = "mV"
                If (m_resolution = 0#) Then m_customUnit = "" ''''decimal case
                
                ''If (m_defreal Like "safe*voltage") Then
                ''''Has been done in auto_eFuse_setReadData_XXX()
                ''''mSV_value = mSV_decimal.Multiply(m_resolution).Add(gD_BaseVoltage)
                
                If (m_defreal = "bincut") Then
                    'm_isBinCut_flag = True
                    ''''<Notice> User maintain
                    ''''-------------------------------------
                    ''''m_catenameVBin = "MS001" ''''<NOTICE> M8 uses MS001 on both power VDD_SOC and VDD_SOC_AON
                    m_catenameVbin = m_catename
                    For Each site In TheExec.sites
                        ''''' Every testing stage needs to check this formula with bincut fused value (2017/5/17, Jack)
                        vbinflag = auto_CheckVddBinInRangeNew(m_catenameVbin, m_resolution)
                        
                        ''''20160329 Add for the offline simulation
                        If (vbinflag = 0 And TheExec.TesterMode = testModeOffline) Then
                            vbinflag = 1
                        End If
                        
                        ''''was m_vddbinEnum, its equal to m_Pmode
                        m_Pmode = VddBinStr2Enum(m_catenameVbin) ''''20160329 add
                        ''tmpVal = auto_VfuseStr_to_Vdd_New(m_bitStrM, m_bitwidth, m_resolution)
                        
                        MaxLevelIndex = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step
                        m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(MaxLevelIndex) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(MaxLevelIndex)
                        m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
                        ''''judge the result
                        If (vbinflag <> 1) Then
                            mSV_value = -999
                            tmpStr = m_catenameVbin + "(Site " + CStr(site) + ") = " + CStr(mSV_value) + " is not in range"
                            TheExec.Datalog.WriteComment tmpStr
                        End If
                        TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsname, scaletype:=m_scale, Unit:=m_unitType, customUnit:=m_customUnit
                    Next site
                ElseIf (m_defreal = "default") Then
                'ElseIf (m_defreal = "safe voltage" Or m_defreal = "decimal") Then
                ''20191202 For Central Meeting
                ''if the field of bincut and safe voltage isn't fused, then the printing will show "0.0000mV"
                    For Each site In TheExec.sites
                        If (mSL_bitSum = 0) Then mSV_value = 0
                    Next
                End If
            
            ElseIf (m_algorithm = "crc") Then
                m_lolmt = 0
                m_hilmt = 0
                m_tsName0 = m_tsname
                ''''20170309, per Jack's recommendation to use Read_Data to do the CRC calculation result (=gS_CFG_CRC_HexStr)
                ''''          and compare with the CRC readCode.
                If (gB_eFuse_Disable_ChkLMT_Flag = True) Then
                    ''''set Pass
                    For Each site In TheExec.sites
                        m_tsname = m_tsName0 + "_" + UCase(mS_Read_calcCRC_hexStr(site))
                        TheExec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0, Tname:=m_tsname
                    Next site
                Else
                    For Each site In TheExec.sites
                        m_crchexStr = UCase(mS_Read_calcCRC_hexStr(site))
                        m_tsname = m_tsName0 + "_" + m_crchexStr
                        ''''<NOTICE>
                        ''''mSV_hexStr  is the CRC HexStr of Read eFuse Category
                        ''''m_crchexStr is the CRC HexStr by the calculation of read bits.
                        If (UCase(mSV_hexStr) = m_crchexStr) Then
                            ''''Pass
                            TheExec.Flow.TestLimit resultVal:=0, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsname
                        Else
                            ''''Fail
                            TheExec.Flow.TestLimit resultVal:=1, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsname
                        End If
                    Next site
                End If
            End If

            If (m_bitwidth >= 32) Then
                ''''keep Limit as Hex String
                ''''process m_lolmt / m_hilmt to hex string with prefix '0x'
                m_lolmt = auto_Value2HexStr(m_lolmt, m_bitwidth)
                m_hilmt = auto_Value2HexStr(m_hilmt, m_bitwidth)
                
                ''''----------------------------------------------
                ''''auto_TestStringLimit compare with lolmt, hilmt
                ''''return  0 means Fail
                ''''return  1 means Pass
                ''''----------------------------------------------
                For Each site In TheExec.sites
                    m_HexStr = mSV_hexStr(site)
                    mSV_value(site) = auto_TestStringLimit(LCase(m_HexStr), CStr(m_lolmt), CStr(m_hilmt)) - 1
                Next site
                ''''mSV_value=0: Pass, = -1 Fail
                m_lolmt = 0
                m_hilmt = 0
                mSV_lolmt = 0
                mSV_hilmt = 0
            Else
                ''''translate to double value
                If (auto_isHexString(CStr(m_lolmt)) = True) Then m_lolmt = auto_HexStr2Value(m_lolmt)
                If (auto_isHexString(CStr(m_hilmt)) = True) Then m_hilmt = auto_HexStr2Value(m_hilmt)
            End If
    
            ''''201812XX update
            If (m_defreal <> "bincut" And m_algorithm <> "crc") Then
                'TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsName, ScaleType:=m_scale, unit:=m_unitType, customUnit:=m_customUnit
                If (m_GetHLlimit = True) Then
                    For Each site In TheExec.sites
                        TheExec.Flow.TestLimit resultVal:=mSV_value(site), lowVal:=mSV_lolmt(site), hiVal:=mSV_hilmt(site), Tname:=m_tsname, scaletype:=m_scale, Unit:=m_unitType, customUnit:=m_customUnit
                    Next site
                Else
                    TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsname, scaletype:=m_scale, Unit:=m_unitType, customUnit:=m_customUnit
                End If
            End If

        End If ''''end of If()
    Next i

    

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Public Function ReverseWave(ByVal InWave As DSPWave, ByRef outWave As DSPWave, ByVal PatBitOrder As String, ByVal SampleSize As Long) As Long
On Error GoTo errHandler
    Dim funcName As String:: funcName = "ReverseWave"
            
    Dim m_tmpArr() As Long
    Dim m_outArr() As Long
    Dim i As Long
    Dim site As Variant
    ReDim m_outArr(SampleSize - 1)
    
                    
    If (PatBitOrder = "bitLast_bit0") Then
        For Each site In TheExec.sites
            m_tmpArr = InWave(site).Data
            For i = 0 To SampleSize - 1
                m_outArr(i) = m_tmpArr(SampleSize - i - 1) ''''save TT
            Next i
            outWave(site).Data = m_outArr
        Next site
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function auto_eFuse_IDS_SetWriteDecimal_ForDsp(ByVal FuseType As eFuseBlockType, _
                                                      ByRef All_Power_data As PinListData, _
                                                      ByRef eFusePower_Pin As PinList, _
                                                      ByVal TestResultFlag As String, _
                                                      Optional showPrint As Boolean = True, _
                                                      Optional chkStage As Boolean = True) As Long
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_IDS_SetWriteDecimal_ForDsp"

    ''----------------------------------------------------
    
    Dim m_len As Long
    Dim i As Long
    Dim m_decimal As Long
    Dim m_bitwidth As Long
    Dim m_resolution As Double
    Dim m_dlogstr As String
    Dim m_stage As String
    Dim m_valStr As String
    Dim ss As Variant
    
    Dim m_catename As String
    Dim index As Long
    Dim mSD_IDSResult As New SiteDouble
    Dim p_ary() As String
    Dim PinCnt As Long
    Dim m_FuseStr As String
    Dim m_tmpStr As String
    Dim m_value As Variant
    Dim m_value_1 As Variant, m_value_2 As Variant
    Dim m_Site As Variant
    Dim Pass_Fail_Flag As New SiteBoolean

    Pass_Fail_Flag = False
    If eFusePower_Pin <> "" Then
        TheExec.DataManager.DecomposePinList eFusePower_Pin, p_ary, PinCnt
    End If
    
    If (FuseType = eFuse_CFG) Then
        m_FuseStr = "CFG"
    End If
    
    For i = 0 To PinCnt - 1
        m_catename = "ids_" + LCase(p_ary(i))
        If (gS_JobName = "cp2") Then m_catename = m_catename & "_85"
        'If (gS_JobName = "cp2") Then m_catename = UCase(Replace(m_catename, "_85", "", 1, 1))
        index = CFGIndex(m_catename)
        If (index <> -1) Then
            m_resolution = CFGFuse.Category(index).Resoultion
            m_bitwidth = CFGFuse.Category(index).BitWidth
            
            If (m_resolution <= 0) Then
                m_tmpStr = funcName + ":: Please have positive resolution " & m_resolution & " (" + m_FuseStr + "::" + m_catename + ")"
                TheExec.AddOutput m_tmpStr
                TheExec.Datalog.WriteComment m_tmpStr
                GoTo errHandler
            End If
            m_resolution = m_resolution * 0.001
            
            For Each m_Site In TheExec.sites
                If TheExec.sites.Item(m_Site).FlagState(TestResultFlag) = logicFalse Then 'Pass
                    Pass_Fail_Flag(m_Site) = True
                Else
                    Pass_Fail_Flag(m_Site) = False
                End If
                mSD_IDSResult(m_Site) = auto_calc_IDS_Decimal(All_Power_data.Pins(p_ary(i)).Value(m_Site), m_resolution)
                If mSD_IDSResult(m_Site) = 0 Then mSD_IDSResult(m_Site) = 1 ' 0 -10mA 0.2-10mA
                If mSD_IDSResult(m_Site) < 0 Then mSD_IDSResult(m_Site) = 0
                If mSD_IDSResult(m_Site) > (2 ^ m_bitwidth - 1) Then
                    mSD_IDSResult(m_Site) = 0
                    Pass_Fail_Flag(m_Site) = False
                End If
                
                If (showPrint) Then
                    m_valStr = Format(All_Power_data.Pins(p_ary(i)).Value(m_Site) * 1000, "0.000000")
                    m_tmpStr = FormatNumeric(m_FuseStr, 4) + "Fuse"
                    m_tmpStr = m_tmpStr + FormatNumeric(" SetWriteVariable_SiteAware ", 10)
                    m_dlogstr = vbTab & "Site(" + CStr(m_Site) + ") " + m_tmpStr + "   " + FormatNumeric(m_catename, 10) + " = " + FormatNumeric(mSD_IDSResult(m_Site), -10) + _
                                " (" + FormatNumeric(m_valStr + " mA", 12) + _
                                " / " + Format(m_resolution * 1000, "0.000000") + "mA)"
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next m_Site
            
            Call auto_eFuse_SetPatTestPass_Flag_SiteAware(m_FuseStr, m_catename, Pass_Fail_Flag, True)
            Call auto_eFuse_SetWriteVariable_SiteAware(m_FuseStr, m_catename, mSD_IDSResult)
        End If
    Next i
    
'    mSD_IDSResult = 0
    
'    m_catename = LCase("IDS_VDD_SOC_Plus_VDD_ISP")
'    Index = CFGIndex(m_catename)
'    m_resolution = CFGFuse.Category(Index).Resoultion
'    m_resolution = m_resolution * 0.001
'    For Each m_Site In TheExec.sites
'        m_value_1 = All_Power_data.Pins("VDD_SOC").Value(m_Site)
'        m_value_2 = All_Power_data.Pins("VDD_ISP").Value(m_Site)
'        m_value = m_value_1 + m_value_2
'        mSD_IDSResult(m_Site) = auto_calc_IDS_Decimal(m_value, m_resolution)
'        If mSD_IDSResult(m_Site) = 0 Then mSD_IDSResult(m_Site) = 1
'        If mSD_IDSResult(m_Site) < 0 Then mSD_IDSResult(m_Site) = 0
'    Next m_Site
'    Call auto_eFuse_SetWriteVariable_SiteAware(m_FuseStr, m_catename, mSD_IDSResult)
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


