Attribute VB_Name = "Module1"
Sub do_assign()
Dim v(6) As Single
Worksheets("assignment").Activate
For i = 3 To 3
    For j = 1 To 7
        For ii = 1 To 6
            v(ii) = Cells(ii + 7, 5).Value
            Next ii
        v(i) = Cells(i + 7, j + 1).Value
        Worksheets("frac_ex").Activate
        For ii = 1 To 6
            Cells(ii + 1, 1).Value = v(ii)
            Next ii
        model
        result = Cells(10, 4).Value
        Worksheets("assignment").Activate
        Cells(i + 20, j + 1).Value = result
        Next j
    Next i
End Sub
Sub demo()
Worksheets("demo").Activate
Randomize

'get params
r = Cells(2, 1).Value
k = Cells(3, 1).Value
vi = Cells(4, 1).Value
ve = Cells(5, 1).Value
f = Cells(6, 1).Value
a = Cells(7, 1).Value

'run a simulation
maxyr = 50
n = k
yr = 0
Cells(10, 1).Value = yr
Cells(10, 2).Value = n
Do
yr = yr + 1
n = Int(n + dndt(n, r, k, vi, ve, f, a) + 0.5)
Cells(10 + yr, 1).Value = yr
Cells(10 + yr, 2).Value = n
Loop Until (n < 1.1) Or (yr >= maxyr)

'if population goes extinct, then blank out remainder of series
If yr < maxyr Then
    For i = yr + 1 To maxyr
        Cells(10 + i, 1).Value = i
        Cells(10 + i, 2).Value = ""
        Next i
End If

End Sub
Sub model()
Worksheets("frac_extinct").Activate
Randomize

'get params
r = Cells(2, 1).Value
k = Cells(3, 1).Value
vi = Cells(4, 1).Value
ve = Cells(5, 1).Value
f = Cells(6, 1).Value
a = Cells(7, 1).Value

'get controls
maxyr = Cells(2, 4).Value
nsim = Cells(3, 4).Value

'run repeated simulations
x = 0
For s = 1 To nsim
    Cells(7, 4).Value = s
    If extinct(maxyr, r, k, vi, ve, f, a) Then x = x + 1
    Next s
frac_extinct = x / nsim

'write results
Cells(10, 4).Value = frac_extinct

End Sub
Function extinct(maxyr, r, k, vi, ve, f, a)
n = k
yr = 0
Do
yr = yr + 1
Cells(6, 4).Value = yr
n = Int(n + dndt(n, r, k, vi, ve, f, a) + 0.5)
Loop Until (n <= 1) Or (yr >= maxyr)
If n < 1 Then extinct = True Else extinct = False
End Function
Function dndt(n, r, k, vi, ve, f, a)
dndt = logistic(n, r, k) + norm(0, n * vi) + norm(0, n * n * ve) + catas(n, f, a)
End Function
Function catas(n, f, a)
If Rnd() < f Then catas = -n * a Else catas = 0
End Function
Function logistic(n, r, k)
logistic = r * n * (k - n) / k
End Function
Function norm(mean, var)
sd = Sqr(var)
norm = mean + sd * z
End Function
Function z()
t = 0
For i = 1 To 12
t = t + Rnd()
Next i
z = t - 6
End Function
