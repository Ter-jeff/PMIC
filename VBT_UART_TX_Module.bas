Attribute VB_Name = "VBT_UART_TX_Module"



Public Function UART_read_response() As Long

    Dim i As Long
    
'    With thehdw.Protocol.ports("UART_PA").NWire.Frames("UART_Snd")
'        .Fields("Data_in").Value = 13   '10
'        .Execute
'    End With
    
    For i = 0 To 200
            
        With TheHdw.Protocol.ports("UART_TX").NWire.Frames("UART_Rcv")
            .Execute tlNWireExecutionType_CaptureInCMEM
        End With

    Next i
  
End Function

Public Function UART_read_response_extended() As Long

    Dim i As Long
    
'    With thehdw.Protocol.ports("UART_PA").NWire.Frames("UART_Snd")
'        .Fields("Data_in").Value = 13   '10
'        .Execute
'    End With
    
    For i = 0 To 1500   '1500
            
        With TheHdw.Protocol.ports("UART_TX").NWire.Frames("UART_Rcv")
            .Execute tlNWireExecutionType_CaptureInCMEM
        End With

    Next i
  
End Function
'

Public Function UART_boot() As Long

    Dim i As Long
    
    For i = 0 To 1000 'TTR-750
            
        With TheHdw.Protocol.ports("UART_TX").NWire.Frames("UART_Rcv")
            .Execute tlNWireExecutionType_CaptureInCMEM
        End With

    Next i
  
End Function
