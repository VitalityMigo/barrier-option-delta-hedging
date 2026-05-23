Attribute VB_Name = "Mod_Utils"
Option Explicit
Option Base 1

Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr

Function mySwitch(TypeOption)
    TypeOption = UCase(TypeOption)
    If TypeOption = "C" Or TypeOption = "CALL" Then
        mySwitch = 1
    ElseIf TypeOption = "P" Or TypeOption = "PUT" Then
        mySwitch = -1
    Else
        mySwitch = "argument has to be call or put"
    End If
End Function

Function mymax(a As Double, b As Double) As Double

If a > b Then
    mymax = a
Else
    mymax = b
End If

End Function

Function myN(x)

myN = WorksheetFunction.Norm_S_Dist(x, True)

End Function

Function Nsd(d As Double) As Double
    Nsd = WorksheetFunction.NormSDist(d)
End Function
