Attribute VB_Name = "VBT_LIB_Common_AP"
Option Explicit
Function VBT_IEDA_Registry(RegistryName As String, Optional OnOff As Boolean = True, Optional DebugPrint As Boolean = True)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "VBT_IEDA_Registry"
    Dim InputStr As String
    If OnOff Then
        Call IEDA_Initialize(InputStr)  'clean up strings
        Call IEDA_GetString(InputStr, RegistryName)  'compose ieda string
        Call IEDA_AutoCheck_Print(InputStr, RegistryName, DebugPrint)   'show log
        Call IEDA_SaveRegistry(InputStr, RegistryName)  'save to registry
    End If

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function
