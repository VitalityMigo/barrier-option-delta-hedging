Attribute VB_Name = "Mod_DeltaHedge"
Option Explicit
Option Base 1

'Load path for common random numbers
Function LoadAllPaths(nbTraj As Long, nbSteps As Long) As Double()

Dim i As Long, j As Long
ReDim allPaths(1 To nbTraj, 0 To nbSteps + 1) As Double

Dim ws As Worksheet
Set ws = Worksheets("MC PATHS")

Dim dataRange As Variant
dataRange = ws.Range(ws.Cells(27, 2), _
            ws.Cells(27 + nbSteps + 1, nbTraj + 1)).Value

For i = 1 To nbTraj
    For j = 0 To nbSteps
        allPaths(i, j) = dataRange(j + 1, i)
    Next j
    allPaths(i, nbSteps + 1) = dataRange(nbSteps + 2, i)
Next i

LoadAllPaths = allPaths

End Function

'Delta hedge function
Function DeltaHedge_UOC(path() As Double, _
                         typeOp As String, _
                         S As Double, _
                         k As Double, _
                         r As Double, _
                         q As Double, _
                         sigma As Double, _
                         t As Double, _
                         H As Double, _
                         nbSteps As Long, _
                         eps As Double) As Double

Dim dt As Double
Dim currentS As Double
Dim currentDelta As Double, prevDelta As Double
Dim currentT As Double
Dim cash As Double
Dim j As Long
Dim initPrice As Double
Dim payoff As Double

dt = t / nbSteps

'Init Price
initPrice = cuo_fran(typeOp, S, k, r, q, sigma, t, H)

'Check if 0
If S >= H Then
    DeltaHedge_UOC = -initPrice
    Exit Function
End If

'Delta init
prevDelta = numdelta_cuo(typeOp, S, k, r, q, sigma, t, H, eps)

'Cash account
cash = initPrice - prevDelta * S

For j = 1 To nbSteps

    'price from path
    currentS = path(j)

    If currentS >= H Then
        'ko
        cash = cash * Exp(r * dt) + prevDelta * currentS
        DeltaHedge_UOC = -cash
        Exit Function
    End If

    'cash cap
    cash = cash * Exp(r * dt)

    currentT = t - j * dt

    If j < nbSteps Then
        'new delta
        currentDelta = numdelta_cuo(typeOp, currentS, k, r, q, sigma, currentT, H, eps)
        cash = cash - (currentDelta - prevDelta) * currentS
        prevDelta = currentDelta
    End If

Next j

'maturity
payoff = WorksheetFunction.Max(currentS - k, 0)

'hedging cost
DeltaHedge_UOC = -(cash + prevDelta * currentS - payoff)

End Function

Function MC_DeltaHedge_UOC(allPaths() As Double, _
                             typeOp As String, _
                             S As Double, _
                             k As Double, _
                             r As Double, _
                             q As Double, _
                             sigma As Double, _
                             t As Double, _
                             H As Double, _
                             nbTraj As Long, _
                             nbSteps As Long, _
                             eps As Double) As Double()

Dim i As Long, j As Long
ReDim hedgeCost(nbTraj) As Double
Dim path() As Double
ReDim path(0 To nbSteps + 1) As Double

For i = 1 To nbTraj
    For j = 0 To nbSteps + 1
        path(j) = allPaths(i, j)
    Next j
    hedgeCost(i) = DeltaHedge_UOC(path, typeOp, S, k, r, q, sigma, t, H, nbSteps, eps)
Next i

MC_DeltaHedge_UOC = hedgeCost

End Function

'Compute base metrics
Function ComputeMetrics(costs() As Double, nbTraj As Long) As Double()

Dim i As Long, k As Long
Dim meanC As Double, stdC As Double
Dim sumC As Double, sumSq As Double
Dim idx95 As Long, idx99 As Long
Dim VaR95 As Double, VaR99 As Double
Dim CVaR95 As Double, CVaR99 As Double
Dim sumTail As Double, countTail As Long
Dim skewnessV As Double, kurtosisV As Double
Dim lossCount As Long, lossPct As Double
Dim sumCube As Double, sumQuad As Double

'mean
sumC = 0
For i = 1 To nbTraj
    sumC = sumC + costs(i)
Next i
meanC = sumC / nbTraj

'std
sumSq = 0
For i = 1 To nbTraj
    sumSq = sumSq + (costs(i) - meanC) ^ 2
Next i
stdC = Sqr(sumSq / (nbTraj - 1))

'VaR 95% et 99% via Large
idx95 = CLng(0.95 * nbTraj)
idx99 = CLng(0.99 * nbTraj)
VaR95 = WorksheetFunction.Large(costs, nbTraj - idx95)
VaR99 = WorksheetFunction.Large(costs, nbTraj - idx99)

'CVaR 95%
sumTail = 0
countTail = nbTraj - idx95
For k = 1 To countTail
    sumTail = sumTail + WorksheetFunction.Large(costs, k)
Next k
If countTail > 0 Then CVaR95 = sumTail / countTail

'CVaR 99%
sumTail = 0
countTail = nbTraj - idx99
For k = 1 To countTail
    sumTail = sumTail + WorksheetFunction.Large(costs, k)
Next k
If countTail > 0 Then CVaR99 = sumTail / countTail

sumCube = 0
sumQuad = 0
lossCount = 0

If stdC > 1E-10 Then
    For i = 1 To nbTraj
        sumCube = sumCube + ((costs(i) - meanC) / stdC) ^ 3
        sumQuad = sumQuad + ((costs(i) - meanC) / stdC) ^ 4
        If costs(i) > 0 Then lossCount = lossCount + 1
    Next i
    skewnessV = sumCube / nbTraj
    kurtosisV = sumQuad / nbTraj - 3
Else
    'guard
    skewnessV = 0
    kurtosisV = 0
    For i = 1 To nbTraj
        If costs(i) > 0 Then lossCount = lossCount + 1
    Next i
End If

lossPct = lossCount / nbTraj

ReDim result(11) As Double
result(1) = meanC
result(2) = stdC
result(3) = VaR95
result(4) = VaR99
result(5) = CVaR95
result(6) = CVaR99
result(7) = skewnessV
result(8) = kurtosisV
result(9) = lossPct
result(10) = WorksheetFunction.Max(costs)
result(11) = WorksheetFunction.Percentile(costs, 0.75) - _
             WorksheetFunction.Percentile(costs, 0.25)

ComputeMetrics = result

End Function

'Hedging metrics
Function ComputeHedgingMetrics(hedgeCost() As Double, _
                                allPaths() As Double, _
                                typeOp As String, _
                                S As Double, _
                                k As Double, _
                                r As Double, _
                                q As Double, _
                                sigma As Double, _
                                t As Double, _
                                H As Double, _
                                nbTraj As Long, _
                                nbSteps As Long, _
                                eps As Double, _
                                initPrice As Double) As Double()

Dim i As Long, j As Long
Dim dt As Double
Dim stepT As Double
Dim nearThreshold As Double

'conditional
Dim sumKO As Double, sumAlive As Double
Dim sumSqKO As Double, sumSqAlive As Double
Dim cntKO As Long, cntAlive As Long
Dim meanKO As Double, meanAlive As Double
Dim stdKO As Double, stdAlive As Double
Dim koRate As Double

'near H
Dim maxDelta As Double
Dim sumRebCost As Double, meanRebCost As Double
Dim sumNearDelta As Double, cntNearBarrier As Long
Dim nearRebCost As Double, pathRebCost As Double
Dim deltaPrev As Double, deltaCurr As Double
Dim rebCost As Double

'others
Dim varDecomp As Double
Dim initDelta As Double
Dim sumStepKO As Long, meanStepKO As Double
Dim jKO As Long

dt = t / nbSteps
nearThreshold = 0.9 * H

'delta init.
initDelta = numdelta_cuo(typeOp, S, k, r, q, sigma, t, H, eps)

For i = 1 To nbTraj

    If allPaths(i, nbSteps + 1) = 0 Then
        sumKO = sumKO + hedgeCost(i)
        cntKO = cntKO + 1

        For jKO = 1 To nbSteps
            If allPaths(i, jKO) >= H Then
                sumStepKO = sumStepKO + jKO
                Exit For
            End If
        Next jKO
    Else
        sumAlive = sumAlive + hedgeCost(i)
        cntAlive = cntAlive + 1
    End If

    deltaPrev = initDelta
    pathRebCost = 0

    For j = 1 To nbSteps - 1
        If allPaths(i, j) >= H Then Exit For
        stepT = t - j * dt
        deltaCurr = numdelta_cuo(typeOp, allPaths(i, j), k, r, q, sigma, stepT, H, eps)
        rebCost = Abs(deltaCurr - deltaPrev) * allPaths(i, j)
        pathRebCost = pathRebCost + rebCost

        'near barrier
        If allPaths(i, j) > nearThreshold Then
            sumNearDelta = sumNearDelta + deltaCurr
            nearRebCost = nearRebCost + rebCost
            cntNearBarrier = cntNearBarrier + 1
        End If

        'max delta
        If Abs(deltaCurr) > maxDelta Then maxDelta = Abs(deltaCurr)

        deltaPrev = deltaCurr
    Next j

    sumRebCost = sumRebCost + pathRebCost

Next i

'conditional mean
If cntKO > 0 Then meanKO = sumKO / cntKO
If cntAlive > 0 Then meanAlive = sumAlive / cntAlive
koRate = cntKO / nbTraj
If cntKO > 0 Then meanStepKO = sumStepKO / cntKO
meanRebCost = sumRebCost / nbTraj
If cntNearBarrier > 0 Then
    sumNearDelta = sumNearDelta / cntNearBarrier
    nearRebCost = nearRebCost / cntNearBarrier
End If

'conditional std
For i = 1 To nbTraj
    If allPaths(i, nbSteps + 1) = 0 Then
        sumSqKO = sumSqKO + (hedgeCost(i) - meanKO) ^ 2
    Else
        sumSqAlive = sumSqAlive + (hedgeCost(i) - meanAlive) ^ 2
    End If
Next i
If cntKO > 1 Then stdKO = Sqr(sumSqKO / (cntKO - 1))
If cntAlive > 1 Then stdAlive = Sqr(sumSqAlive / (cntAlive - 1))

'decomposition
varDecomp = (stdKO ^ 2) * koRate + (stdAlive ^ 2) * (1 - koRate)

Dim sumC2 As Double
sumC2 = 0
For i = 1 To nbTraj
    sumC2 = sumC2 + hedgeCost(i)
Next i

'data
ReDim result(14) As Double
result(1) = initDelta
result(2) = meanKO
result(3) = meanAlive
result(4) = stdKO
result(5) = stdAlive
result(6) = koRate
result(7) = cntKO
result(8) = meanStepKO
result(9) = maxDelta
result(10) = meanRebCost
result(11) = sumNearDelta
result(12) = nearRebCost
result(13) = varDecomp
result(14) = (sumC2 / nbTraj) / initPrice

ComputeHedgingMetrics = result

End Function

Sub RunDeltaHedge_UOC()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Worksheets("DYN. HEDGING").Activate

Dim typeOp As String
Dim S As Double, k As Double, r As Double, q As Double
Dim sigma As Double, t As Double, H As Double, eps As Double
Dim nbTraj As Long, nbSteps As Long

typeOp = Range("C4").Value
S = CDbl(Range("C5").Value)
k = CDbl(Range("C6").Value)
H = CDbl(Range("C7").Value)
r = CDbl(Range("C8").Value)
q = CDbl(Range("C9").Value)
sigma = CDbl(Range("C10").Value)
t = CDbl(Range("C11").Value)
nbTraj = CLng(Range("C12").Value)
nbSteps = CLng(Range("C13").Value)
eps = CDbl(Range("C14").Value)

'theo pricing
Dim initPrice As Double
initPrice = cuo_fran(typeOp, S, k, r, q, sigma, t, H)

'load path for CRN
Dim allPaths() As Double
allPaths = LoadAllPaths(nbTraj, nbSteps)

'sim
Dim hedgeCost() As Double
hedgeCost = MC_DeltaHedge_UOC(allPaths, typeOp, S, k, r, q, sigma, t, H, nbTraj, nbSteps, eps)

Dim met() As Double
Dim hmet() As Double
met = ComputeMetrics(hedgeCost, nbTraj)
hmet = ComputeHedgingMetrics(hedgeCost, allPaths, typeOp, S, k, r, q, sigma, t, H, nbTraj, nbSteps, eps, initPrice)

'hull ratio
Dim hullRatio As Double
If initPrice <> 0 Then hullRatio = met(2) / initPrice

Dim runTime As Long
runTime = 0

'key results
Worksheets("DYN. HEDGING").Activate

Range("C18").Value = initPrice
Range("C19").Value = met(1)
Range("C20").Value = met(2)
Range("C21").Value = hullRatio
Range("C22").Value = met(3)
Range("C23").Value = met(4)
Range("C24").Value = met(5)
Range("C25").Value = met(6)


Range("G18").Value = met(7)
Range("G19").Value = met(8)
Range("G20").Value = met(9)
Range("G21").Value = met(10)
Range("G22").Value = met(11)



Range("G4").Value = hmet(1)
Range("G5").Value = hmet(2)
Range("G6").Value = hmet(3)
Range("G7").Value = hmet(4)
Range("G8").Value = hmet(5)
Range("G9").Value = hmet(9)
Range("G10").Value = hmet(11)
Range("G11").Value = hmet(13)
Range("G12").Value = hmet(10) / nbSteps
Range("G13").Value = hmet(12)
If hmet(10) <> 0 Then
    Range("G14").Value = hmet(12) / (hmet(10) / nbSteps)
Else
    Range("G14").Value = 0
End If

'histo data
Dim i As Long
Range("B29").Value = "Hedging Cost"
For i = 1 To nbTraj
    Cells(i + 29, 2).Value = hedgeCost(i)
Next i

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

MsgBox "Simulation completed ! " & _
       "Price UOC = " & Round(initPrice, 4) & Chr(13) & _
       "Hull Ratio = " & Round(hullRatio, 4), _
       vbInformation, "Delta-Hedging UOC"

End Sub

