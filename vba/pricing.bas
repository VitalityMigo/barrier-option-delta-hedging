Attribute VB_Name = "mod_pricing"
Option Explicit
Option Base 1

Function BS(TypeOption, S, k, r, q, sigma, t)
Dim d1 As Double, d2 As Double, z As Integer
d1 = (Log(S / k) + (r - q + 0.5 * sigma ^ 2) * t) / _
    (sigma * Sqr(t))
d2 = d1 - sigma * Sqr(t)
z = mySwitch(TypeOption)
BS = z * (S * Exp(-q * t) * Nsd(z * d1) - _
    k * Exp(-r * t) * Nsd(z * d2))
End Function


Function myBSprice(typ As String, _
                    S As Double, _
                    k As Double, _
                    r As Double, _
                    div As Double, _
                    sigma As Double, _
                    t As Double)
'A generic black-Scholes pricer for European call and put paying divident yield div

Dim d1 As Double, d2 As Double
Dim z As Integer

z = mySwitch(typ)

d1 = (Log(S / k) + (r - div + 0.5 * sigma ^ 2) * t) / (sigma * Sqr(t))
d2 = d1 - sigma * Sqr(t)
myBSprice = z * S * Exp(-div * t) * myN(z * d1) _
    - z * k * Exp(-r * t) * myN(z * d2)

End Function

Function myDelta(typ As String, _
                    S As Double, _
                    k As Double, _
                    r As Double, _
                    div As Double, _
                    sigma As Double, _
                    t As Double)
Dim d1 As Double, d2 As Double
Dim z As Integer

z = mySwitch(typ)

d1 = (Log(S / k) + (r - div + 0.5 * sigma ^ 2) * t) / (sigma * Sqr(t))

myDelta = Exp(-div * t) * z * WorksheetFunction.Norm_S_Dist(z * d1, True)

End Function

Function myGamma(S, k, r, div, sigma, t)
Dim d1 As Double
d1 = (Log(S / k) + (r - div + 0.5 * sigma ^ 2) * t) / _
    (sigma * Sqr(t))
myGamma = WorksheetFunction.Norm_S_Dist(d1, False) * Exp(-div * t) / _
    (S * sigma * Sqr(t))

End Function

Function mytheta(TypeOption, S, k, r, div, sigma, t)
Dim d1 As Double, d2 As Double
Dim z As Integer
    d1 = (Log(S / k) + (r - div + 0.5 * sigma ^ 2) * t) / _
    (sigma * Sqr(t))
    d2 = d1 - sigma * Sqr(t)
    z = mySwitch(TypeOption)
mytheta = -S * WorksheetFunction.NormDist(d1, 0, 1, False) * sigma * Exp(-div * t) / _
    (2 * Sqr(t)) + z * div * S * myN(z * d1) * Exp(-div * t) - _
    z * r * k * Exp(-r * t) * myN(z * d2)
End Function

Function myVega(S, k, r, div, sigma, t)
Dim d1 As Double
d1 = (Log(S / k) + (r - div + 0.5 * sigma ^ 2) * t) / _
    (sigma * Sqr(t))
myVega = S * Sqr(t) * WorksheetFunction.NormDist(d1, 0, 1, False) * Exp(-div * t)
End Function

Function myRho(TypeOption, S, k, r, div, sigma, t)
Dim d1 As Double, d2 As Double
Dim z As Integer

Function cuo(TypeOption, S, k, r, q, sigma, t, H)
    d1 = (Log(S / k) + (r - q + 0.5 * sigma ^ 2) * t) / _
    (sigma * Sqr(t))
    d2 = d1 - sigma * Sqr(t)
    lam = (r - q + (sigma ^ 2) / 2) / sigma ^ 2
    x1 = Log(S / H) / (sigma * Sqr(t)) + lam * sigma * Sqr(t)
    y1 = Log(H / S) / (sigma * Sqr(t)) + lam * sigma * Sqr(t)
    y = Log(H ^ 2 / (S * k)) / (sigma * Sqr(t)) + lam * sigma * Sqr(t)
cui = S * Nsd(x1) * Exp(-q * t) - k * Exp(-r * t) * Nsd(x1 - sigma * Sqr(t)) - _
    S * Exp(-q * t) * ((H / S) ^ (2 * lam)) * (Nsd(-y) - Nsd(-y1)) + _
    k * Exp(-r * t) * ((H / S) ^ (2 * lam - 2)) * (Nsd(-y + sigma * Sqr(t)) - Nsd(-y1 + sigma * Sqr(t)))
    cuo = BS(TypeOption, S, k, r, q, sigma, t) - cui
End Function

Function cuo_fran(TypeOption, S, k, r, q, sigma, t, H)
    Dim a As Double, b As Double
    Dim z1 As Double, z2 As Double, z3 As Double, z4 As Double
    Dim z5 As Double, z6 As Double, z7 As Double, z8 As Double
    
    a = (H / S) ^ ((2 * (r - q)) / (sigma ^ 2) - 1)
    b = (H / S) ^ ((2 * (r - q)) / (sigma ^ 2) + 1)
    
    z1 = (Log(S / k) + (r - q + (sigma ^ 2) / 2) * t) / (sigma * Sqr(t))
    z2 = z1 - sigma * Sqr(t)
    z3 = (Log(S / H) + (r - q + (sigma ^ 2) / 2) * t) / (sigma * Sqr(t))
    z4 = z3 - sigma * Sqr(t)
    z5 = (Log(S / H) - (r - q - (sigma ^ 2) / 2) * t) / (sigma * Sqr(t))
    z6 = z5 - sigma * Sqr(t)
    z7 = (Log(S * k / H ^ 2) - (r - q - (sigma ^ 2) / 2) * t) / (sigma * Sqr(t))
    z8 = z7 - sigma * Sqr(t)
   
    cuo_fran = S * Exp(-q * t) * (Nsd(z1) - Nsd(z3) - b * (Nsd(z6) - Nsd(z8))) - _
        k * Exp(-r * t) * (Nsd(z2) - Nsd(z4) - a * (Nsd(z5) - Nsd(z7)))
    
End Function

Function numdelta_cuo(TypeOption, S, k, r, q, sigma, t, H, eps)

numdelta_cuo = (cuo_fran(TypeOption, S + eps, k, r, q, sigma, t, H) - _
    cuo_fran(TypeOption, S - eps, k, r, q, sigma, t, H)) / (2 * eps)
    
End Function


Function cashop(TypeOption, S, k, r, q, sigma, t)
    d1 = (Log(S / k) + (r - q + 0.5 * sigma ^ 2) * t) / _
    (sigma * Sqr(t))
    d2 = d1 - sigma * Sqr(t)
    cashop = Exp(-r * t) * Nsd(Switch(TypeOption) * d2)
End Function

Function ndel_cash(TypeOption, S, k, r, q, sigma, t, eps)

ndel_cash = (cashop(TypeOption, S + eps, k, r, q, sigma, t) - _
    cashop(TypeOption, S - eps, k, r, q, sigma, t)) / (2 * eps)
    
End Function
