Attribute VB_Name = "VBT_UART_RX_Module"
Public Function UART_write_pmgr() As Long

    SendCmdOnly "pmgr bincut-check disable"
    
End Function
'
'Public Function UART_write_run_sc_11() As Long
'
'    SendCmdOnly "sc run 11"
'
'End Function
