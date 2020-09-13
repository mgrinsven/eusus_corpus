Attribute VB_Name = "Module1"
Option Explicit
Dim p As Double, f As Double, s As Double, n As Double, r As Double

Public Function SplineFit(TableName As Range, term As Double) As Double
'This is a translation (into VBA) of a Fortran 77 routine contained in Press et al.,
'"Numerical Recipes". The pertinent chapter is on line (as of 1/24/03) at the address
'http://lib-www.lanl.gov/numerical/bookfpdf/f3-3.pdf . Note that I am using here what
'many authors call "natural splines". This means I assume the second derivative is zero
'at the boundaries. Also see http://mathworld.wolfram.com/CubicSpline.html for a good
'discussion of the mathematics behind this technique.
'
'TableName is a named range of cells that can be used as a lookup table.

Dim n As Integer, i As Integer, k As Integer, khi As Integer, _
    klo As Integer, ncols As Integer
Dim x(20) As Double, y(20) As Double, y2(20) As Double, u(20) As Double
Dim p As Double, qn As Double, sig As Double, un As Double, a As Double, _
    b As Double, h As Double
n = TableName.Rows.Count
For i = 1 To n
    x(i) = TableName.Cells(i, 1).Value
Next i
If term < x(1) Or term > x(n) Then
    SplineFit = CVErr(xlErrValue)
    Return
End If
For i = 1 To n
    y(i) = TableName.Cells(i, 2)
Next
y2(1) = 0#
u(1) = 0
For i = 2 To n - 1
    sig = (x(i) - x(i - 1)) / (x(i + 1) - x(i - 1))
    p = sig * y2(i - 1) + 2#
    y2(i) = (sig - 1#) / p
    u(i) = (6# * ((y(i + 1) - y(i)) / (x(i + 1) - x(i)) - (y(i) - y(i - 1)) _
            / (x(i) - x(i - 1))) / (x(i + 1) - x(i - 1)) - sig * u(i - 1)) / p
Next i
qn = 0#
un = 0#
y2(n) = (un - qn * u(n - 1)) / (qn * y2(n - 1) + 1#)
For k = n - 1 To 1 Step -1
    y2(k) = y2(k) * y2(k + 1) + u(k)
Next k
klo = 1
khi = n
While (khi - klo > 1)
    k = (khi + klo) / 2
    If (x(k) > term) Then
    khi = k
    Else
        klo = k
    End If
Wend
h = x(khi) - x(klo)
a = (x(khi) - term) / h
b = (term - x(klo)) / h
SplineFit = a * y(klo) + b * y(khi) + ((a ^ 3 - a) * y2(klo) + (b ^ 3 - b) * y2(khi)) * (h ^ 2) / 6
End Function

Public Function APR(p As Double, s As Double, f As Double, r As Double, n As Double) As Double
Application.Volatile
Dim a As Double, delt As Double, g As Double, gprime As Double
r = r / 12
a = r * 1.05
Do
  g = (a * (1 + a) ^ n) / (((1 + a) ^ n) - 1) - (p + f - s) * r * (1 + r) ^ n / _
      ((p - s) * (((1 + r) ^ n) - 1))
  gprime = ((a + 1) ^ (n - 1) * ((a + 1) ^ (n + 1) - a * (n + 1) - 1)) / _
           ((a + 1) ^ n - 1) ^ 2
  delt = g / gprime
  a = a - delt
Loop Until Abs(delt) <= 0.00001
APR = a * 12
End Function
