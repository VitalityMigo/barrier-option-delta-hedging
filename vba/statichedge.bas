Attribute VB_Name = "Mod_StaticHedge"
Option Base 1

Function Stat_hedge_cuo(typeOp, S, k, r, q, sigma, t, H, n)

Dim dt As Double
Dim i As Long, j As Long
Dim datev As Double, Ptf_val As Double, BS_new As Double

dt = t / (n - 1)
ReDim NB_Van(n) As Double
ReDim BS_date(n) As Double
ReDim Str_Van(n) As Double
ReDim Ech_Van(n) As Double

For i = 1 To n
    NB_Van(i) = 0
    BS_date(i) = 0
Next i

NB_Van(1) = 1
Str_Van(1) = k
Ech_Van(1) = t

For i = 2 To n
    datev = t - (i - 1) * dt
    Str_Van(i) = H
    Ech_Van(i) = t + (2 - i) * dt
    For j = 1 To i - 1
        BS_date(j) = BS(typeOp, H, Str_Van(j), r, q, sigma, Ech_Van(j) - datev)
    Next j
    Ptf_val = WorksheetFunction.SumProduct(NB_Van, BS_date)
    BS_new = BS(typeOp, H, Str_Van(i), r, q, sigma, Ech_Van(i) - datev)
    NB_Van(i) = -Ptf_val / BS_new
Next i

For j = 1 To n
    BS_date(j) = BS(typeOp, S, Str_Van(j), r, q, sigma, Ech_Van(j))
Next j
Stat_hedge_cuo = WorksheetFunction.SumProduct(NB_Van, BS_date)

End Function

'change from original code: common random numbers from MC Paths, and accounts for premium
'to compare to dynamic hedge

Sub RunStaticHedge_UOC()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Worksheets("STATIC HEDGING").Activate

Dim typeOp As String
Dim S As Double, k As Double, r As Double, q As Double
Dim sigma As Double, t As Double, H As Double, eps As Double
Dim nbTraj As Long, nbSteps As Long, nbVan As Long
Dim convPrice As Double
Dim convRow As Long
Dim nV As Integer

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
nbVan = CLng(Range("C15").Value)

'theo price
Dim initPriceUOC As Double
initPriceUOC = cuo_fran(typeOp, S, k, r, q, sigma, t, H)

'static portfolio
Dim dtV As Double
dtV = t / (nbVan - 1)
ReDim NB_Van(nbVan) As Double
ReDim Str_Van(nbVan) As Double
ReDim Ech_Van(nbVan) As Double
ReDim BS_date(nbVan) As Double

Dim ii As Long, jj As Long
For ii = 1 To nbVan
    NB_Van(ii) = 0
    BS_date(ii) = 0
Next ii

NB_Van(1) = 1
Str_Van(1) = k
Ech_Van(1) = t

Dim datev As Double, Ptf_val As Double, BS_new As Double
For ii = 2 To nbVan
    datev = t - (ii - 1) * dtV
    Str_Van(ii) = H
    Ech_Van(ii) = t + (2 - ii) * dtV
    For jj = 1 To ii - 1
        BS_date(jj) = BS(typeOp, H, Str_Van(jj), r, q, sigma, Ech_Van(jj) - datev)
    Next jj
    Ptf_val = WorksheetFunction.SumProduct(NB_Van, BS_date)
    BS_new = BS(typeOp, H, Str_Van(ii), r, q, sigma, Ech_Van(ii) - datev)
    NB_Van(ii) = -Ptf_val / BS_new
Next ii

'inital value
For jj = 1 To nbVan
    BS_date(jj) = BS(typeOp, S, Str_Van(jj), r, q, sigma, Ech_Van(jj))
Next jj
Dim initValSH As Double
initValSH = WorksheetFunction.SumProduct(NB_Van, BS_date)

'rep. error
Dim replError As Double
replError = initValSH - initPriceUOC

Dim replErrorPct As Double
If initPriceUOC <> 0 Then replErrorPct = replError / initPriceUOC

'paths
Dim allPaths() As Double
allPaths = LoadAllPaths(nbTraj, nbSteps)

'modified hedge cost
Dim hedgeCost() As Double
ReDim hedgeCost(nbTraj) As Double
Dim i As Long
Dim currentS As Double, activeFlag As Integer
Dim finalVal As Double, payoff As Double

For i = 1 To nbTraj
    currentS = allPaths(i, nbSteps)
    activeFlag = CInt(allPaths(i, nbSteps + 1))
    If activeFlag = 0 Then
        finalVal = 0
    Else
        For jj = 1 To nbVan
            BS_date(jj) = WorksheetFunction.Max(currentS - Str_Van(jj), 0)
        Next jj
        finalVal = WorksheetFunction.SumProduct(NB_Van, BS_date)
    End If
    payoff = activeFlag * WorksheetFunction.Max(currentS - k, 0)
    hedgeCost(i) = initValSH - finalVal + payoff - initPriceUOC
Next i

'cond. metrics
Dim sumKO As Double, sumAlive As Double
Dim cntKO As Long, cntAlive As Long
Dim meanKO As Double, meanAlive As Double
Dim koRate As Double

For i = 1 To nbTraj
    If allPaths(i, nbSteps + 1) = 0 Then
        sumKO = sumKO + hedgeCost(i)
        cntKO = cntKO + 1
    Else
        sumAlive = sumAlive + hedgeCost(i)
        cntAlive = cntAlive + 1
    End If
Next i

If cntKO > 0 Then meanKO = sumKO / cntKO
If cntAlive > 0 Then meanAlive = sumAlive / cntAlive
koRate = cntKO / nbTraj

'init. delta
Dim initDelta As Double
initDelta = numdelta_cuo(typeOp, S, k, r, q, sigma, t, H, eps)

'write
Worksheets("STATIC HEDGING").Activate

Range("C18").Value = initPriceUOC
Range("C19").Value = initValSH
Range("C20").Value = hedgeCost(1)
Range("C21").Value = replErrorPct

Range("G4").Value = initDelta
Range("G5").Value = meanKO
Range("G6").Value = meanAlive

Range("B29").Value = "Hedging Cost"
For i = 1 To nbTraj
    Cells(i + 29, 2).Value = hedgeCost(i)
Next i

'convergence analysis
convRow = 30
For nV = 5 To 300 Step 5
    convPrice = Stat_hedge_cuo(typeOp, S, k, r, q, sigma, t, H, nV)
    Cells(convRow, 6).Value = nV
    Cells(convRow, 7).Value = convPrice
    Cells(convRow, 8).Value = convPrice - initPriceUOC
    convRow = convRow + 1
Next nV

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

MsgBox "Static Hedge completed !" & Chr(13) & _
       "UOC Price      = " & Round(initPriceUOC, 4) & Chr(13) & _
       "Static Price   = " & Round(initValSH, 4) & Chr(13) & _
       "Repl. Error %  = " & Round(replErrorPct * 100, 2) & "%", _
       vbInformation, "Static Hedging UOC"

End Sub

