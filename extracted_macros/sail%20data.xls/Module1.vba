Attribute VB_Name = "Module1"



Rem  Lines starting with "rem" are remarks only.
Rem  You should'nt need to change anything on this page, be careful if you do!

Rem  DISPLACEMENT / LWL RATIO
Function dldoc(x, x1, x2, x3, x4)
Attribute dldoc.VB_ProcData.VB_Invoke_Func = " \n14"

Rem  The functon DLDOC calculates the degree of compatibility (doc) between the data and the optimum.
Rem  All doc's range from 0 to 1, 1 being perfect.
Rem  When the function dldoc(x) is called, "x" should be the disp/length ratio for that data set
Rem  The distribution is a trapezoid, starting at a minimum value,
Rem  (x1), ramping up to a flat (x2, x3) and dropping back to zero at x4.

dldoc = 0
If x > x1 Then
dldoc = (x - x1) / (x2 - x1)
End If
If x > x2 Then
dldoc = 1
End If
If x > x3 Then
dldoc = (x - x4) / (x3 - x4)
End If
If x > x4 Then
dldoc = 0
End If
End Function


Rem  TED BREWER'S COMFORT FACTOR
Function cfdoc(x, x1, x2, x3, x4)
Attribute cfdoc.VB_ProcData.VB_Invoke_Func = " \n14"

Rem  When the function cfdoc is called "x" should be the comfort factor for that data set
Rem  X2 - X3 is range of the optimum comfort factor, with a trapezoid distribution, as above.
Rem  X1 is the minimum value
Rem  X4 is the maximum value

cfdoc = 0
If x > x1 Then
cfdoc = (x - x1) / (x2 - x1)
End If
If x > x2 Then
cfdoc = 1
End If
If x > x3 Then
cfdoc = (x - x4) / (x3 - x4)
End If
If x > x4 Then
cfdoc = 0
End If
End Function


Rem  CAPSIZE RISK
Function crdoc(x, x3, x4)
Attribute crdoc.VB_ProcData.VB_Invoke_Func = " \n14"

Rem  When the function crdoc is called "x" should be the capsize risk for that data set.
Rem  X0 is the optimum capsize ratio.
Rem  This is a linear distribution, dropping from 1 at x3 to 0 at x4.

crdoc = 1
If x > x3 Then
crdoc = (x - x4) / (x3 - x4)
End If
If x > x4 Then
crdoc = 0
End If
End Function


Rem  SAIL AREA / DISPLACEMENT RATIO
Function sddoc(x, x1, x2, x3, x4)
Attribute sddoc.VB_ProcData.VB_Invoke_Func = " \n14"

Rem  When the function sddoc is called "x" should be the sail area/disp for that data set.
Rem  Another trapezoidal distribution.

sddoc = 0
If x > x1 Then
sddoc = (x - x1) / (x2 - x1)
End If
If x > x2 Then
sddoc = 1
End If
If x > x3 Then
sddoc = (x - x4) / (x3 - x4)
End If
If x > x4 Then
sddoc = 0
End If
End Function


Rem  VMVH DOC
Function vmvhdoc(x, x1, x2, x3, x4)
Attribute vmvhdoc.VB_ProcData.VB_Invoke_Func = " \n14"

Rem  When the function vmvhdoc is called "x" should be the Vm/Vh for that data set.
Rem  This is a trapezoid distribution.
Rem  Rewards boats with enough sail area and low weight to move faster than hull speed.

vmvhdoc = 0
If x > x1 Then
vmvhdoc = (x - x1) / (x2 - x1)
End If
If x > x2 Then
vmvhdoc = 1
End If
If x > x3 Then
vmvhdoc = (x - x4) / (x3 - x4)
End If
If x > x4 Then
vmvhdoc = 0
End If

End Function


Rem  Maximum radial acceleration experienced by a sailor sleeping 1.5' inboard
Rem  of the maximum beam during 10 degree rolls.
Rem Function aceldoc(x, x1, x2, x3, x4)

Function aceldoc(x, x3, x4)

Rem  When the function loadoc is called "x" should be the acceleration for that data set.
Rem  Trapezoid distribution.

Rem changed to linear dist., low accel is always better, like capsize risk

Rem  Used to filter out boats that respond to fast.

Rem aceldoc = 0
Rem If x > x1 Then
Rem aceldoc = (x - x1) / (x2 - x1)
Rem End If
Rem If x > x2 Then
Rem aceldoc = 1
Rem End If
Rem If x > x3 Then
Rem aceldoc = (x - x4) / (x3 - x4)
Rem End If
Rem If x > x4 Then
Rem aceldoc = 0

aceldoc = 1
If x > x3 Then
aceldoc = (x - x4) / (x3 - x4)
End If
If x > x4 Then
aceldoc = 0
End If
End Function


Function lbdoc(x, x1, x2, x3, x4)
Attribute lbdoc.VB_ProcData.VB_Invoke_Func = " \n14"

Rem  When the function lbdoc is called, "x" should be the L/B for that data set.
Rem  Trapezoid distribution.
Rem  Used to filter out boats that are too narrow or too fat.

lbdoc = 0
If x > x1 Then
lbdoc = (x - x1) / (x2 - x1)
End If
If x > x2 Then
lbdoc = 1
End If
If x > x3 Then
lbdoc = (x - x4) / (x3 - x4)
End If
If x > x4 Then
lbdoc = 0
End If
End Function

Function hedge(x)

Rem x is the value set by the spinner that is used to divide the standard deviation.
Rem this function links the value to a HEDGE.
hedge = "SOMEWHAT CLOSE"
If x < 2.1 Then
hedge = "CLOSE"
End If
If x < 1.1 Then
hedge = "VERY CLOSE"
End If
End Function














