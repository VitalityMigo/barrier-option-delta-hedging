Attribute VB_Name = "Mod_Comp"
Option Explicit
Option Base 1

Sub RunComparison()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Worksheets("MC PATHS").Activate

Dim typeOp As String
Dim S As Double, k As Double, r As Double, q As Double
Dim sigma As Double, T_base As Double, H As Double, eps As Double
Dim nbTraj As Long, nbSteps As Long
Dim dhKORate As Double

typeOp = Range("C4").Value
S = CDbl(Range("C5").Value)
k = CDbl(Range("C6").Value)
H = CDbl(Range("C7").Value)
r = CDbl(Range("C8").Value)
q = CDbl(Range("C9").Value)
sigma = CDbl(Range("C10").Value)
T_base = CDbl(Range("C11").Value)
nbTraj = CLng(Range("C12").Value)
nbSteps = CLng(Range("C13").Value)
eps = CDbl(Range("C14").Value)
dhKORate = CDbl(Range("C20").Value)

Dim dhHullRatio As Double
Dim dhMeanCost As Double

Worksheets("DYN. HEDGING").Activate
dhHullRatio = CDbl(Range("C21").Value)
dhMeanCost = CDbl(Range("C19").Value)

'theo px
Dim initPriceUOC As Double
initPriceUOC = cuo_fran(typeOp, S, k, r, q, sigma, T_base, H)

'SH at Various freq.
Dim nVanArray(3) As Integer
nVanArray(1) = 10
nVanArray(2) = 50
nVanArray(3) = 100

Dim shPrice(3) As Double
Dim shReplErrorPct(3) As Double
Dim n As Integer

For n = 1 To 3
    shPrice(n) = Stat_hedge_cuo(typeOp, S, k, r, q, sigma, T_base, H, nVanArray(n))
    If initPriceUOC <> 0 Then
        shReplErrorPct(n) = (shPrice(n) - initPriceUOC) / initPriceUOC
    End If
Next n

Worksheets("COMPARISON").Activate

'hull
Cells(5, 3).Value = dhHullRatio
Cells(5, 4).Value = 0
Cells(5, 5).Value = 0
Cells(5, 6).Value = 0

'cost
Cells(6, 3).Value = dhMeanCost
Cells(6, 4).Value = shPrice(1) - initPriceUOC
Cells(6, 5).Value = shPrice(2) - initPriceUOC
Cells(6, 6).Value = shPrice(3) - initPriceUOC

'rep error
Dim dhMeanCostPct As Double
If initPriceUOC <> 0 Then dhMeanCostPct = dhMeanCost / initPriceUOC
Cells(7, 3).Value = Abs(dhMeanCost / initPriceUOC)
Cells(7, 4).Value = Abs((initPriceUOC - shPrice(1)) / initPriceUOC)
Cells(7, 5).Value = Abs((initPriceUOC - shPrice(2)) / initPriceUOC)
Cells(7, 6).Value = Abs((initPriceUOC - shPrice(3)) / initPriceUOC)

'reb cost
Cells(8, 3).Value = CDbl(Worksheets("DYN. HEDGING").Range("G12").Value)
Cells(8, 4).Value = 0
Cells(8, 5).Value = 0
Cells(8, 6).Value = 0

'KO rate
Cells(9, 3).Value = dhKORate
Cells(9, 4).Value = dhKORate
Cells(9, 5).Value = dhKORate
Cells(9, 6).Value = dhKORate

'MATRIX SCENARIO
Dim moneyness(4) As Double
moneyness(1) = 0.8
moneyness(2) = 0.9
moneyness(3) = 1#
moneyness(4) = 1.1

Dim maturities(3) As Double
maturities(1) = 0.25
maturities(2) = 0.5
maturities(3) = 0.75

Dim nMoney As Integer, nMat As Integer
nMoney = 4
nMat = 3

Dim topRows(3) As Integer
topRows(1) = 13
topRows(2) = 20
topRows(3) = 27

Dim wsSce As Worksheet
Set wsSce = Worksheets("DYN. HEDGING (Cont'd)")

Dim wsComp As Worksheet
Set wsComp = Worksheets("COMPARISON")

Dim m As Integer, t As Integer
Dim Scen As Double, Tcen As Double
Dim dhCost As Double, shCostN As Double
Dim initPriceScen As Double

For n = 1 To 3
    For m = 1 To nMoney
        For t = 1 To nMat

            Scen = moneyness(m) * k
            Tcen = maturities(t)

            dhCost = CDbl(wsSce.Cells(m + 11, t + 2).Value)
            initPriceScen = cuo_fran(typeOp, Scen, k, r, q, sigma, Tcen, H)
            shCostN = Stat_hedge_cuo(typeOp, Scen, k, r, q, sigma, Tcen, H, nVanArray(n)) - initPriceScen
            wsComp.Cells(topRows(n) + m - 1, t + 2).Value = dhCost - shCostN

        Next t
    Next m
Next n

'3 PATH ANALYSIS
Dim nVanPath As Integer
nVanPath = 50

Dim allPaths() As Double
allPaths = LoadAllPaths(nbTraj, nbSteps)

'SH
Dim dtV As Double
dtV = T_base / (nVanPath - 1)
ReDim NB_Van(nVanPath) As Double
ReDim Str_Van(nVanPath) As Double
ReDim Ech_Van(nVanPath) As Double
ReDim BS_date(nVanPath) As Double

Dim ii As Long, jj As Long
For ii = 1 To nVanPath
    NB_Van(ii) = 0
    BS_date(ii) = 0
Next ii

NB_Van(1) = 1
Str_Van(1) = k
Ech_Van(1) = T_base

Dim datev As Double, Ptf_val As Double, BS_new As Double
For ii = 2 To nVanPath
    datev = T_base - (ii - 1) * dtV
    Str_Van(ii) = H
    Ech_Van(ii) = T_base + (2 - ii) * dtV
    For jj = 1 To ii - 1
        BS_date(jj) = BS(typeOp, H, Str_Van(jj), r, q, sigma, Ech_Van(jj) - datev)
    Next jj
    Ptf_val = WorksheetFunction.SumProduct(NB_Van, BS_date)
    BS_new = BS(typeOp, H, Str_Van(ii), r, q, sigma, Ech_Van(ii) - datev)
    NB_Van(ii) = -Ptf_val / BS_new
Next ii

For jj = 1 To nVanPath
    BS_date(jj) = BS(typeOp, S, Str_Van(jj), r, q, sigma, Ech_Van(jj))
Next jj
Dim initValSH50 As Double
initValSH50 = WorksheetFunction.SumProduct(NB_Van, BS_date)

'path selection
Dim pathFar As Long, pathNear As Long, pathKO As Long
Dim maxSFar As Double, maxSNear As Double
Dim maxSi As Double
Dim i As Long, j As Long

pathFar = 0
pathNear = 0
pathKO = 0
maxSFar = 999
maxSNear = 0

For i = 1 To nbTraj
    maxSi = 0
    For j = 1 To nbSteps
        If allPaths(i, j) > maxSi Then maxSi = allPaths(i, j)
    Next j

    'KO
    If allPaths(i, nbSteps + 1) = 0 And pathKO = 0 Then
        pathKO = i
    End If

    'Far
    If allPaths(i, nbSteps + 1) = 1 And maxSi < 0.85 * H Then
        If maxSi < maxSFar Then
            maxSFar = maxSi
            pathFar = i
        End If
    End If

    'Near H
    If allPaths(i, nbSteps + 1) = 1 And maxSi > 0.95 * H Then
        If maxSi > maxSNear Then
            maxSNear = maxSi
            pathNear = i
        End If
    End If
Next i

'Stats
Dim pathIdx(3) As Long
pathIdx(1) = pathFar
pathIdx(2) = pathNear
pathIdx(3) = pathKO

Dim dhCostPath As Double, shCostPath As Double
Dim finalS As Double, maxS As Double
Dim maxDeltaPath As Double
Dim activeF As Integer
Dim currentS As Double, payoff As Double, finalVal As Double

'format
Dim startRow As Long
startRow = 11

Dim p As Integer
For p = 1 To 3

    If pathIdx(p) = 0 Then GoTo NextPath

    'stats
    finalS = allPaths(pathIdx(p), nbSteps)
    activeF = CInt(allPaths(pathIdx(p), nbSteps + 1))

    maxS = 0
    For j = 1 To nbSteps
        If allPaths(pathIdx(p), j) > maxS Then maxS = allPaths(pathIdx(p), j)
    Next j

    'DH cost
    Dim pathArr() As Double
    ReDim pathArr(0 To nbSteps + 1) As Double
    For j = 0 To nbSteps + 1
        pathArr(j) = allPaths(pathIdx(p), j)
    Next j
    dhCostPath = DeltaHedge_UOC(pathArr, typeOp, S, k, r, q, sigma, T_base, H, nbSteps, eps)

    'SH cost
    currentS = allPaths(pathIdx(p), nbSteps)
    activeF = CInt(allPaths(pathIdx(p), nbSteps + 1))
    If activeF = 0 Then
        finalVal = 0
    Else
        For jj = 1 To nVanPath
            BS_date(jj) = WorksheetFunction.Max(currentS - Str_Van(jj), 0)
        Next jj
        finalVal = WorksheetFunction.SumProduct(NB_Van, BS_date)
    End If
    payoff = activeF * WorksheetFunction.Max(currentS - k, 0)
    shCostPath = initValSH50 - finalVal + payoff - initPriceUOC

    'max delta
    maxDeltaPath = 0
    Dim stepT As Double
    Dim dt As Double
    dt = T_base / nbSteps
    For j = 1 To nbSteps - 1
        If allPaths(pathIdx(p), j) >= H Then Exit For
        stepT = T_base - j * dt
        Dim deltaJ As Double
        deltaJ = Abs(numdelta_cuo(typeOp, allPaths(pathIdx(p), j), k, r, q, sigma, stepT, H, eps))
        If deltaJ > maxDeltaPath Then maxDeltaPath = deltaJ
    Next j

    wsComp.Cells(startRow + 1, 9).Value = finalS
    wsComp.Cells(startRow + 2, 9).Value = maxS
    wsComp.Cells(startRow + 3, 9).Value = dhCostPath
    wsComp.Cells(startRow + 4, 9).Value = shCostPath
    wsComp.Cells(startRow + 5, 9).Value = dhCostPath - shCostPath
    wsComp.Cells(startRow + 6, 9).Value = maxDeltaPath
    
    startRow = startRow + 8

NextPath:
Next p

'Delta Path Analysis 2

Dim dtPath As Double
dtPath = T_base / nbSteps

Dim jWrite As Long
Dim stepTWrite As Double

For jWrite = 0 To nbSteps

    stepTWrite = T_base - jWrite * dtPath
    wsComp.Cells(jWrite + 39, 2).Value = jWrite

    'Far
    If pathFar > 0 Then
        wsComp.Cells(jWrite + 39, 3).Value = allPaths(pathFar, jWrite)
        If jWrite < nbSteps And allPaths(pathFar, jWrite) < H Then
            wsComp.Cells(jWrite + 39, 4).Value = numdelta_cuo(typeOp, allPaths(pathFar, jWrite), k, r, q, sigma, stepTWrite, H, eps)
        End If
    End If

    'Near
    If pathNear > 0 Then
        wsComp.Cells(jWrite + 39, 5).Value = allPaths(pathNear, jWrite)
        If jWrite < nbSteps And allPaths(pathNear, jWrite) < H Then
            wsComp.Cells(jWrite + 39, 6).Value = numdelta_cuo(typeOp, allPaths(pathNear, jWrite), k, r, q, sigma, stepTWrite, H, eps)
        End If
    End If

    'KO
    If pathKO > 0 Then
        wsComp.Cells(jWrite + 39, 7).Value = allPaths(pathKO, jWrite)
        If jWrite < nbSteps And allPaths(pathKO, jWrite) < H Then
            wsComp.Cells(jWrite + 39, 8).Value = numdelta_cuo(typeOp, allPaths(pathKO, jWrite), k, r, q, sigma, stepTWrite, H, eps)
        End If
    End If

Next jWrite

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

MsgBox "Comparison Completed !", vbInformation, "RunComparison"

End Sub
