Attribute VB_Name = "Module1"
Option Explicit

Public Const A1  As Double = 0.3265
Public Const A2  As Double = -1.07
Public Const A3  As Double = -0.5339
Public Const A4  As Double = 0.01569
Public Const A5  As Double = -0.05165
Public Const A6  As Double = 0.5475
Public Const A7  As Double = -0.7361
Public Const A8  As Double = 0.1844
Public Const A9  As Double = 0.1056
Public Const A10 As Double = 0.6134
Public Const A11 As Double = 0.721
    
Public Function pC7Calc(gCell, tCell)

    Dim g As Double
    Dim B As Double
    
    Dim gi As Double
    Dim gi2 As Double
    g = Range("gC7Plus").Value
    gi = 1# / g
    gi2 = gi * gi
    
    B = Range("tBoilC7Plus").Value
    pC7Calc = Exp((8.3634 - 0.0566 * gi) - 0.001 * (0.24244 + 2.2898 * gi + 0.11857 * gi2) * B + 0.0000001 * (1.4685 + 3.648 * gi + 0.47227 * gi2) * B ^ 2 - 0.0000000001 * (0.42019 + 1.6977 * gi2) * B ^ 3)
    
End Function

Public Static Function tC7Calc(gCell, tCell)

    Dim g As Double
    Dim B As Double
    
    g = Range("gC7Plus").Value
    B = Range("tBoilC7Plus").Value

    tC7Calc = 341.7 + 811 * g + (0.4244 + 0.1174 * g) * B + (0.4669 - 3.2623 * g) * 100000# / B

End Function

Public Function gHCOnly(gTotal, fH2S, fCO2, Mair)

gHCOnly = (gTotal - (fCO2 * Worksheets("Constant Gas Properties").Cells(23, 4).Value + fH2S * Worksheets("Constant Gas Properties").Cells(24, 4).Value) / Mair) / (1 - fCO2 - fH2S)
End Function



Public Function zCalc(pPR, TPR)

    Dim zNew As Double
    Dim zLast As Double
    Dim rhoPR As Double
    Dim TPRI As Double
    Dim A As Variant
    Dim kount As Integer
    
    Dim dfdz As Double
    Dim fofz As Double
    Dim dfofz As Double
    Dim dz As Double
    
    Dim c1 As Double
    Dim c2 As Double
    Dim c3 As Double
    Dim c4 As Double
    
    Const zTol As Double = 0.00001
    Const kountMax As Integer = 50
    Const zMin As Double = 0.01
    Const zMax As Double = 10#
    
    TPRI = 1 / TPR
    
    zLast = 1
    rhoPR = 0.27 * pPR / (zLast * TPR)

    c1 = A1 + A2 * TPRI + A3 * TPRI ^ 3 + A4 * TPRI ^ 4 + A5 * TPRI ^ 5
    c2 = A6 + A7 * TPRI + A8 * TPRI ^ 2
    c3 = A9 * (A7 * TPRI + A8 * TPRI ^ 2)
    c4 = A10 * (1 + A11 * rhoPR ^ 2) * (rhoPR ^ 2 * TPRI ^ 3) * Exp(-A11 * rhoPR ^ 2)
    zCalc = 1 + c1 * rhoPR + c2 * rhoPR ^ 2 - c3 * rhoPR ^ 5 + c4
    kount = 0
    Do
        zNew = 1 + c1 * rhoPR + c2 * rhoPR ^ 2 - c3 * rhoPR ^ 5 + c4
        'MsgBox zNew & "    " & zLast
        fofz = zLast - zNew
        dfdz = (1 + c1 * rhoPR / zNew _
                + 2 * c2 * rhoPR ^ 2 / zNew _
                - 5 * c3 * rhoPR ^ 5 / zNew _
                + 2 * A10 * rhoPR ^ 2 * TPRI ^ 3 / zNew _
                * (1 + A11 * rhoPR ^ 2 - (A11 * rhoPR ^ 2) ^ 2) * Exp(-A11 * rhoPR ^ 2))
        dz = -fofz / dfdz
        'MsgBox dfdz & "    " & dz
        zLast = zLast + dz
        If (zLast < zMin) Then zLast = zMin
        If (zLast > zMax) Then zLast = zMax
        rhoPR = 0.27 * pPR / (zLast * TPR)
        c4 = A10 * (1 + A11 * rhoPR ^ 2) * (rhoPR ^ 2 * TPRI ^ 3) * Exp(-A11 * rhoPR ^ 2)
        kount = kount + 1
      Loop Until (Abs(zLast - zNew) <= zTol Or (kount > kountMax))
    zCalc = zLast
    If (zCalc = zMin Or zCalc = zMax Or kount > kountMax) Then zCalc = "@NA"
End Function

Public Function muCalc(MW, T, rho)
    
    Dim A As Double
    Dim B As Double
    Dim C As Double
    Dim rhoGCC As Double
    
    rhoGCC = rho * 453.6 / (12 * 2.54) ^ 3
    
    A = ((9.379 + 0.01607 * MW) * T ^ 1.5) / (209.2 + 19.62 * MW + T)
    B = 3.448 + 986.4 / T + 0.01009 * MW
    C = 2.447 - 0.2224 * B
    
    muCalc = 0.0001 * A * Exp(B * rhoGCC ^ C)
    
End Function

Public Function cgCalc(pPR, TPR, z)

    Dim dzdp As Double
    Dim TPRI As Double
    Dim rhoPR As Double
    
    TPRI = 1 / TPR
    rhoPR = 0.27 * pPR / (z * TPR)
    
    dzdp = A1 + A2 * TPRI + A3 * (TPRI ^ 3) + A4 * (TPRI ^ 4) + A5 * (TPRI ^ 5) + 2 * (A6 + A7 * TPRI + A8 * (TPRI ^ 2)) * rhoPR - 5 * A9 * TPRI * (A7 + A8 * TPRI) * (rhoPR ^ 4) + 2 * A10 * rhoPR * (TPRI ^ 3) * (1 + A11 * (rhoPR ^ 2) - A11 * (rhoPR ^ 4)) * Exp(-A11 * (rhoPR ^ 2))
    cgCalc = 1 / pPR - 0.27 / (z * z * TPR) * (dzdp / (1 + rhoPR * dzdp / z))

End Function

Public Function tCGasCalc(gSurf, gResvr)

If (gResvr > 0) Then
    tCGasCalc = 187 + 330 * gResvr - 71.5 * gResvr ^ 2
Else
    If (gSurf > 0) Then
        tCGasCalc = 168 + 325 * gSurf - 12.5 * gSurf ^ 2
    Else
        tCGasCalc = "@NA"
    End If
End If

End Function

Public Function pCGasCalc(gSurf, gResvr)

If (gResvr <> 0) Then
    pCGasCalc = 706 - 51.7 * gResvr - 11.1 * gResvr ^ 2
Else
    If (gSurf > 0) Then
        pCGasCalc = 677 + 15 * gSurf - 37.5 * gSurf ^ 2
    Else
        pCGasCalc = "@NA"
    End If
End If

End Function

Public Function Const1(TRI)
    Const1 = A1 + A2 * TRI + A3 * TRI ^ 3 + A4 * TRI ^ 4 + A5 * TRI ^ 5
End Function


Public Function Const2(TRI)
    Const2 = A6 + A7 * TRI + A8 * TRI ^ 2
End Function

Public Sub Const3(TRI)
    Const3 = A9 * (A7 * TRI + A8 * TRI ^ 2)
End Sub


Public Function Const4(TRI, rhoPR)
    Const4 = A10 * (1 + A11 * rhoPR ^ 2) * (rhoPR ^ 2 * TRI ^ 3) * Exp(-A11 * rhoPR ^ 2)
End Function


Public Function z2Phase(TPR, pPR)
Const c0 As Double = 2.24353
Const c1 As Double = -0.0375281
Const c2 As Double = -3.56539
Const c3 As Double = 0.000829231
Const c4 As Double = 1.53428
Const c5 As Double = 0.131987

z2Phase = c0 + c1 * pPR + c2 / TPR + c3 * pPR ^ 2 + c4 / TPR ^ 2 + c5 * (pPR / TPR)

End Function
