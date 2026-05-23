Attribute VB_Name = "Mod_Scenarios"
Option Explicit
Option Base 1

Sub RunScenarios_UOC()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Randomize

'params
Worksheets("MC PATHS").Activate

Dim typeOp As String
Dim k As Double, H As Double, r As Double
Dim q As Double, sigma As Double, eps As Double
Dim nbVan As Long

typeOp = Range("C4").Value
k = CDbl(Range("C6").Value)
H = CDbl(Range("C7").Value)
r = CDbl(Range("C8").Value)
q = CDbl(Range("C9").Value)
sigma = CDbl(Range("C10").Value)
eps = CDbl(Range("C14").Value)
nbVan = CLng(Range("C15").Value)

Dim wsOut As Worksheet
Set wsOut = Worksheets("DYN. HEDGING (Cont'd)")

'from dyn. hedging sheet
Dim nbTraj As Long, nbSteps As Long
nbTraj = CLng(wsOut.Range("C4").Value)
nbSteps = CLng(wsOut.Range("C5").Value)

'scenarios
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

wsOut.Activate

Dim rowIdx As Integer
rowIdx = 20

Dim m As Integer, tIdx As Integer, i As Long
Dim S As Double, t As Double
Dim initPrice As Double
Dim hullRatioDH As Double
Dim metDH() As Double
Dim costsDH() As Double

'sim
Dim dtDH As Double, intermDH As Double, volSqrtDH As Double
Dim cashDH As Double
Dim prevDelta As Double, currDelta As Double
Dim currentT As Double
Dim payoffDH As Double
Dim currentS As Double, activeFlag As Integer
Dim finalVal As Double, payoff As Double
Dim jPath As Long, kPath As Long, jDH As Long
Dim active As Integer
Dim path() As Double

'delta path
Dim bestPath() As Double
Dim bestMaxS As Double, maxSi As Double
Dim bestT As Double
Dim hasBestPath As Boolean
hasBestPath = False
bestMaxS = 0

For m = 1 To nMoney
    For tIdx = 1 To nMat

        S = moneyness(m) * k
        t = maturities(tIdx)

        'theoretical price
        initPrice = cuo_fran(typeOp, S, k, r, q, sigma, t, H)

        'guard
        If initPrice <= 0 Then
            wsOut.Cells(rowIdx, 7).Value = moneyness(m)
            wsOut.Cells(rowIdx, 8).Value = t
            wsOut.Cells(rowIdx, 9).Value = 0
            rowIdx = rowIdx + 1
            GoTo NextScenario
        End If
        
        'DH constants
        dtDH = t / nbSteps
        intermDH = (r - q - 0.5 * sigma ^ 2) * dtDH
        volSqrtDH = sigma * Sqr(dtDH)

        'hedging sim
        ReDim costsDH(nbTraj) As Double
        ReDim path(0 To nbSteps + 1) As Double

        For i = 1 To nbTraj

            'common path
            path(0) = S
            active = 1

            For jPath = 1 To nbSteps
                path(jPath) = path(jPath - 1) * Exp(intermDH + volSqrtDH * WorksheetFunction.Norm_S_Inv(Rnd()))
                If path(jPath) >= H Then
                    active = 0
                    For kPath = jPath + 1 To nbSteps
                        path(kPath) = H
                    Next kPath
                    Exit For
                End If
            Next jPath
            path(nbSteps + 1) = active

            'near barrier path
        If active = 1 And moneyness(m) = 1# Then
                maxSi = 0
                Dim jMax As Long
                For jMax = 1 To nbSteps
                    If path(jMax) > maxSi Then maxSi = path(jMax)
                Next jMax
            If maxSi > bestMaxS Then
                bestMaxS = maxSi
                bestT = t
                ReDim bestPath(0 To nbSteps + 1) As Double
                Dim jCopy As Long
                For jCopy = 0 To nbSteps + 1
                    bestPath(jCopy) = path(jCopy)
            Next jCopy
            hasBestPath = True
        End If
    End If

            'DH
            prevDelta = numdelta_cuo(typeOp, S, k, r, q, sigma, t, H, eps)
            cashDH = initPrice - prevDelta * S

            For jDH = 1 To nbSteps
                currentS = path(jDH)
                If currentS >= H Then
                    cashDH = cashDH * Exp(r * dtDH) + prevDelta * currentS
                    costsDH(i) = -cashDH
                    GoTo NextPathDH
                End If
                cashDH = cashDH * Exp(r * dtDH)
                currentT = t - jDH * dtDH
                If jDH < nbSteps Then
                    currDelta = numdelta_cuo(typeOp, currentS, k, r, q, sigma, currentT, H, eps)
                    cashDH = cashDH - (currDelta - prevDelta) * currentS
                    prevDelta = currDelta
                End If
            Next jDH
            payoffDH = WorksheetFunction.Max(currentS - k, 0)
            costsDH(i) = -(cashDH + prevDelta * currentS - payoffDH)

NextPathDH:

        Next i

        'stats
        metDH = ComputeMetrics(costsDH, nbTraj)

        If initPrice <> 0 Then
            hullRatioDH = metDH(2) / initPrice
        End If

        wsOut.Cells(m + 11, tIdx + 2).Value = metDH(1)
        wsOut.Cells(m + 4, tIdx + 7).Value = hullRatioDH
        wsOut.Cells(m + 11, tIdx + 7).Value = metDH(3)

        'Scenarios DH
        wsOut.Cells(rowIdx, 7).Value = moneyness(m)
        wsOut.Cells(rowIdx, 8).Value = t
        wsOut.Cells(rowIdx, 9).Value = initPrice
        wsOut.Cells(rowIdx, 10).Value = metDH(1)
        wsOut.Cells(rowIdx, 11).Value = metDH(2)
        wsOut.Cells(rowIdx, 12).Value = hullRatioDH
        wsOut.Cells(rowIdx, 13).Value = metDH(3)
        wsOut.Cells(rowIdx, 14).Value = metDH(4)
        wsOut.Cells(rowIdx, 15).Value = metDH(5)
        wsOut.Cells(rowIdx, 16).Value = metDH(6)

        rowIdx = rowIdx + 1

NextScenario:
    Next tIdx
Next m

'write
If hasBestPath Then
    Dim jWrite As Long
    Dim stepTWrite As Double
    Dim dtWrite As Double
    dtWrite = bestT / nbSteps

    For jWrite = 0 To nbSteps
        wsOut.Cells(jWrite + 20, 2).Value = jWrite
        wsOut.Cells(jWrite + 20, 3).Value = bestPath(jWrite)
        If jWrite < nbSteps Then
            stepTWrite = bestT - jWrite * dtWrite
            wsOut.Cells(jWrite + 20, 4).Value = numdelta_cuo(typeOp, bestPath(jWrite), k, r, q, sigma, stepTWrite, H, eps)
        End If
    Next jWrite
End If

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

MsgBox "Scenarios simulated ! " & (rowIdx - 20) & " scenarios calculated.", _
       vbInformation, "RunScenarios_UOC"

End Sub

