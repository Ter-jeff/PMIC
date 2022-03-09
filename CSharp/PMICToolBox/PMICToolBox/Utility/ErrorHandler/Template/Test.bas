Attribute VB_Name = "VBT_Module"
Option Explicit

' This module should be used for VBT Tests.  All functions in this module
' will be available to be used from the Test Instance sheet.
' Additional modules may be added as needed (all starting with "VBT_").
'
' The required signature for a VBT Test is:
'
' Public Function FuncName(<arglist>) As Long
'   where <arglist> is any list of arguments supported by VBT Tests.
'
' See online help for supported argument types in VBT Tests.
'
'
' It is highly suggested to use error handlers in VBT Tests.  A sample
' VBT Test with a suggeseted error handler is shown below:
'
' Function FuncName() As Long
'     On Error GoTo errHandler
'
'     Exit Function
' errHandler:
'     If AbortTest Then Exit Function Else Resume Next
' End Function

Public Function Test1() As Long
' Public Function FuncName(<arglist>) As Long
'   where <arglist> is any list of arguments supported by VBT Tests.
'
x = thehdw.PPMU.Pins("PinY").Read()
End Function


Public Sub Test2() As Long

x = thehdw.PPMU.Pins("PinY").Read()

End Sub


Public Function FIMV() As Long
    On Error GoTo errHandler

'Define a site aware variable to store the measured voltage
 Dim x As New PinListData
 
'Program ports pins to connect PPMU. Hint: call thehdw.PPMU.pins().connect
 Call thehdw.PPMU.Pins("PinY").Connect
 
'Program ports to gate on PPMU. Hint: thehdw.PPMU.Pins().Gate=
 thehdw.PPMU.Pins("PinY").Gate = tlOn

'And also force 100uA current using PPMU to ports. Hint: call thehdw.PPMU.Pins().ForceI()
 Call thehdw.PPMU.Pins("PinY").ForceI(100 * uA)

'Wait for 1ms. Hint: call thehdw.wait()
 Call thehdw.Wait(1 * ms)
 
'Read the measured voltage back to the site aware variable. Hint:thehdw.PPMU.Pins().Read()
 x = thehdw.PPMU.Pins("PinY").Read()
 
'Judge the test mode, if test mode is offline, give the value to 0.7. Hint: theexec.testermode
 If TheExec.TesterMode = testModeOffline Then
     x.AddPin ("vcc")
     x.Pins("vcc").Value(0) = 0.7
 End If
'Multiple the measured voltage by 2. Hint: PLD.math.multiply()
 x.Math.Multiply (2)
'Compare the test result to test limit in flow. Hint: theexec.flow.testlimit
 TheExec.Flow.TestLimit x.Pins("vcc"), 1, 1.5
 

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error Occured in FVMI test instance!"
End Function

