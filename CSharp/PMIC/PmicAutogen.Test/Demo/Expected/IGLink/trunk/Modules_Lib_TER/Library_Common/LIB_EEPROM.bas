Attribute VB_Name = "LIB_EEPROM"
Option Explicit

Type tDIB_EEPROM
    bIsEEPROMinfoOK As Boolean
    
    bPrintEEPROMInfoAlready As Boolean
    
    sHasEEPROM As String
    sIsProgrammed As String
    sPartNum As String
    sSerialNum As String
End Type

Public g_DIB_EEPROM As tDIB_EEPROM


'======================================================
' ___SOP__________________________________________
'   1. put ReadEEPROM      into OnProgramValidated
'   2. put PrintEEPROMInfo into OnProgramStarted
'======================================================

Public Function ReadEEPROM()    'grab eeprom info to global variable
    On Error GoTo ErrorHandler
        
    '--- reset variable ---
    g_DIB_EEPROM.bPrintEEPROMInfoAlready = False    'refresh flag
    g_DIB_EEPROM.bIsEEPROMinfoOK = False
    g_DIB_EEPROM.sHasEEPROM = "EEPROM not existence!"
    g_DIB_EEPROM.sIsProgrammed = "EEPROM unprogramed!"
    g_DIB_EEPROM.sPartNum = "N/A"
    g_DIB_EEPROM.sSerialNum = "N/A"
        
    '--- main ---
    If TheHdw.Digital.Calibration.DIB.HasEEPROM = True Then
        g_DIB_EEPROM.sHasEEPROM = "EEPROM existence!"
        
        If TheHdw.DIB.eeprom.IsProgrammed = True Then
            g_DIB_EEPROM.bIsEEPROMinfoOK = True     'eeprom info OK
            
            g_DIB_EEPROM.sIsProgrammed = "EEPROM programed!"
            
            g_DIB_EEPROM.sPartNum = TheHdw.DIB.eeprom.PartNumber
            g_DIB_EEPROM.sSerialNum = TheHdw.DIB.eeprom.SerialNumber
        End If
    End If

Exit Function

ErrorHandler:
    LIB_ErrorDescription ("ReadEEPROM")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PrintEEPROMInfo()   'show eeprom info to datalog
    On Error GoTo ErrorHandler
    
    If g_DIB_EEPROM.bPrintEEPROMInfoAlready = False Then    'print eeprom info once only
        If g_DIB_EEPROM.bIsEEPROMinfoOK = True Then
            'Call TheExec.Flow.TestLimit(1, 1, 1, , , , , , g_DIB_EEPROM., , , , , , , tlForceNone)
            TheExec.Datalog.WriteComment "TeradyneSerialsID = " & g_DIB_EEPROM.sSerialNum
            TheExec.Datalog.WriteComment "TeradynePartID = " & g_DIB_EEPROM.sPartNum
        Else
            TheExec.Datalog.WriteComment g_DIB_EEPROM.sHasEEPROM
            TheExec.Datalog.WriteComment g_DIB_EEPROM.sIsProgrammed
            TheExec.Datalog.WriteComment "TeradyneSerialsID = N/A"
            TheExec.Datalog.WriteComment "TeradynePartID = N/A"
        End If
        
        g_DIB_EEPROM.bPrintEEPROMInfoAlready = True
    End If
    
Exit Function

ErrorHandler:
    LIB_ErrorDescription ("PrintEEPROMInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function
