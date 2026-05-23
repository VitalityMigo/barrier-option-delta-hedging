Attribute VB_Name = "Mod_simulation"

Option Base 1


Function GeneratePath(S As Double, _
                      r As Double, _
                      q As Double, _
                      sigma As Double, _
                      t As Double, _
                      H As Double, _
                      nbSteps As Long) As Double()

Dim dt As Double
Dim interm As Double, volSqrt As Double
Dim j As Long, k As Long
Dim activeFlag As Integer

dt = t / nbSteps
interm = (r - q - 0.5 * sigma ^ 2) * dt
volSqrt = sigma * Sqr(dt)

ReDim path(0 To nbSteps + 1) As Double
path(0) = S
activeFlag = 1

For j = 1 To nbSteps
    path(j) = path(j - 1) * Exp(interm + volSqrt * _
              WorksheetFunction.Norm_S_Inv(Rnd()))
    If path(j) >= H Then
        activeFlag = 0
        For k = j + 1 To nbSteps
            path(k) = H
        Next k
        Exit For
    End If
Next j

path(nbSteps + 1) = activeFlag
GeneratePath = path

End Function

Function GenerateAllPaths(S As Double, _
                          r As Double, _
                          q As Double, _
                          sigma As Double, _
                          t As Double, _
                          H As Double, _
                          nbTraj As Long, _
                          nbSteps As Long) As Double()

Dim i As Long, j As Long
Dim path() As Double

ReDim allPaths(1 To nbTraj, 0 To nbSteps + 1) As Double

For i = 1 To nbTraj
    path = GeneratePath(S, r, q, sigma, t, H, nbSteps)
    For j = 0 To nbSteps + 1
        allPaths(i, j) = path(j)
    Next j
Next i

GenerateAllPaths = allPaths

End Function

Sub RunSimulation()

Randomize
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'params
Worksheets("MC PATHS").Activate

Dim S As Double, r As Double, q As Double
Dim sigma As Double, t As Double, H As Double
Dim nbTraj As Long, nbSteps As Long

S = CDbl(Range("C5").Value)
H = CDbl(Range("C7").Value)
r = CDbl(Range("C8").Value)
q = CDbl(Range("C9").Value)
sigma = CDbl(Range("C10").Value)
t = CDbl(Range("C11").Value)
nbTraj = CLng(Range("C12").Value)
nbSteps = CLng(Range("C13").Value)

'paths
Dim allPaths() As Double
allPaths = GenerateAllPaths(S, r, q, sigma, t, H, nbTraj, nbSteps)

'write
Dim i As Long, j As Long
Cells(26, 1).Value = "Step"
For i = 1 To nbTraj
    Cells(26, i + 1).Value = "Path " & i
Next i

For j = 0 To nbSteps
    Cells(j + 27, 1).Value = j
    For i = 1 To nbTraj
        Cells(j + 27, i + 1).Value = allPaths(i, j)
    Next i
Next j

'active vs. KO
Dim activeRow As Long
activeRow = 27 + nbSteps + 1
Cells(activeRow, 1).Value = "Active"
For i = 1 To nbTraj
    Cells(activeRow, i + 1).Value = allPaths(i, nbSteps + 1)
Next i

Dim koCount As Long, aliveCount As Long
For i = 1 To nbTraj
    If allPaths(i, nbSteps + 1) = 0 Then
        koCount = koCount + 1
    Else
        aliveCount = aliveCount + 1
    End If
Next i

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

MsgBox nbTraj & " paths generated (" & nbSteps & " steps)." & Chr(13) & _
       "Knockouts : " & koCount & Chr(13) & _
       "Alive : " & aliveCount, _
       vbInformation, "Simulation"

End Sub



