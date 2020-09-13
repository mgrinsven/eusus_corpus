Attribute VB_Name = "Random"
'distribution types
Global Const DISTFIXED = 0
Global Const DISTUNIFORM = 1
Global Const DISTEXP = 2
Global Const DISTNORMTRUNC = 3
Global Const DISTPERT = 4   'beta variable with given mode and std. dev. = range/6
Global Const DISTBETA = 5   'beta variable with given mean and std. dev. = range /6
Global Const DISTGAMMA = 6 'not implemented yet

'probability distributions
Type ProbDist
    DistType As Integer
    seed As Single
    param1 As Single
    param2 As Single
    param3 As Single
    param4 As Single
End Type

Global ProbDistGlobal As ProbDist

Function ExpectedValue(pdist As ProbDist) As Single
ExpectedValue = 0
Select Case pdist.DistType
    Case DISTFIXED
        ExpectedValue = pdist.param1
    Case DISTUNIFORM
        ExpectedValue = (pdist.param1 + pdist.param2) / 2
    Case DISTEXP
        ExpectedValue = pdist.param1
    Case DISTNORMTRUNC
        ExpectedValue = pdist.param1    'truncation not implemented yet
    Case DISTPERT
        ExpectedValue = (pdist.param1 + 4 * pdist.param3 + pdist.param2) / 6
    Case Else
End Select
End Function

Function RandomBeta(a As Double, b As Double, nextrand As Single) As Double
'  BETA VARIATE GENERATOR                             */
'  From Algorithm due to R.C.H. Cheng                 */
'  Reference: pg 309 of Bratley, Fox, and Schrage,    */
' A GUIDE TO SIMULATION, 2nd Ed. Springer-Verlag     */
'*******************************************************/
' variable definitions */
Dim u1 As Double
Dim u2 As Double
Dim v As Double
Dim w As Double
Static con(4) As Double

con(1) = b
If (a < b) Then con(1) = a
If (con(1) <= 1#) Then
    con(1) = 1# / con(1)
Else
    con(1) = Sqr((a + b - 2#) / (2# * a * b - a - b))
End If
con(2) = a + b
con(3) = a + 1# / con(1)
u1 = nextrand
u2 = Rnd(-u1)
again:
   u1 = Rnd(-u2)
   u2 = Rnd(-u1)
   v = con(1) * Log(u1 / (1# - u1))
   w = a * Exp(v)
If (((con(2)) * Log((con(2)) / (b + w)) + (con(3)) * v - 1.3862944) < (Log(u1 * u1 * u2))) Then
    GoTo again
Else
   RandomBeta = (w / (b + w))
End If

End Function

Function RandomNext(RandomPrev As Single) As Single
RandomPrev = Rnd(-RandomPrev)       'update seed
RandomNext = RandomPrev
End Function

Function RandomNormal(m As Double, s As Double, nextrand As Single) As Double
' NORMAL VARIATE GENERATOR                                        */
'  Box-Muller Technique mean m and standard deviation = s          */
'  Reference: pg 46 of Lewis and Orav, SIMULATION METHODOLOGY FOR  */
' STATISTICIANS, OPERATIONS ANALYSIS, and ENGINEERS, Wadsworth.   */
'********************************************************************/
' variable definitions */
Dim u1 As Double
Dim u2 As Double
Dim X As Double

u1 = Rnd(-nextrand)
u2 = Rnd(-u1)
X = Sqr(-2# * Log(u1)) * Cos(2# * (3.14159265) * u2)
RandomNormal = (m + X * s)

End Function

Function RandomPoisson(Expectedcount As Single) As Integer
Dim sumexponentials As Single
sumexponentials = 0
Dim nextrand As Single
Dim sumcount As Integer
sumcount = -1
While sumexponentials <= Expectedcount
    sumcount = sumcount + 1
    nextrand = RandomNext(ProbDistGlobal.seed)
    sumexponentials = sumexponentials - Log(nextrand)
Wend
RandomPoisson = sumcount
End Function
Function RandomBernoulli(probtrue As Single, newseed As Single) As Integer
Dim p As Single
Dim nextrand As Single
nextrand = RandomNext(newseed)
p = nextrand
If p <= probtrue Then RandomBernoulli = True Else RandomBernoulli = False
End Function
Function RandomVariate(pdist As ProbDist) As Single
'get next random variate in series from pdist
Dim rtime As Single
Dim mean As Double
Dim stddev As Double
Dim shape1 As Double
Dim shape2 As Double
Dim low As Single
Dim high As Single
Dim mode As Single
Dim nextrand As Single

nextrand = RandomNext(pdist.seed)

Select Case pdist.DistType
    Case DISTFIXED
        rtime = pdist.param1
    Case DISTUNIFORM
        rtime = pdist.param1 + nextrand * (pdist.param2 - pdist.param1)
    Case DISTEXP
        If nextrand < SIMTINY / 100 Then nextrand = SIMTINY / 100
        rtime = -Log(nextrand) * pdist.param1
    Case DISTNORMTRUNC
        mean = pdist.param1
        stddev = pdist.param2
        'this routine is biased high: need to rewrite to compensate for truncation
        Dim found As Integer
        If stddev > SIMTINY Then
            found = False
        Else
            found = True
            rtime = mean
        End If
        While Not found
            rtime = RandomNormal(mean, stddev, nextrand)
            If rtime > SIMTINY Then
                found = True
            Else
                nextrand = RandomNext(pdist.seed)
            End If
        Wend
    Case DISTPERT
        low = pdist.param1
        high = pdist.param2
        mode = pdist.param3
        'estimate mean using PERT weights
        mean = (low + 4 * mode + high) / 6  'PERT estimate of mean
        meanfrac = (mean - low) / (high - low)  'potential division by zero; check data first
        'now determine shape parameters for Beta distn on support (0,1) assuming mean is known and std. dev. is 1/6. i.e. 6*sigma equals support
        shape2 = (1 - meanfrac) * (1 - meanfrac) * meanfrac * 36 - (1 - meanfrac)
        shape1 = meanfrac / (1 - meanfrac) * shape2 'potential division by zero; check data first
        'get random beta variate on support (0,1)
        rtime = RandomBeta(shape1, shape2, nextrand)
        'convert to random beta variate on support (low,high)
        rtime = low + rtime * (high - low)
    Case DISTBETA
        low = pdist.param1
        high = pdist.param2
        mean = pdist.param3
        meanfrac = (mean - low) / (high - low)  'potential division by zero; check data first
        'now determine shape parameters for Beta distn on support (0,1) assuming mean is known and std. dev. is 1/6. i.e. 6*sigma equals support
        shape2 = (1 - meanfrac) * (1 - meanfrac) * meanfrac * 36 - (1 - meanfrac)
        shape1 = meanfrac / (1 - meanfrac) * shape2 'potential division by zero; check data first
        'get random beta variate on support (0,1)
        rtime = RandomBeta(shape1, shape2, nextrand)
        'convert to random beta variate on support (low,high)
        rtime = low + rtime * (high - low)
    Case DISTGAMMA
        'not implemented yet
    Case Else
        MsgBox "Unrecognized probability distribution "
End Select

rtime = CLng(rtime * 1000) / 1000   'truncate nuisance digits
If rtime < SIMTINY Then rtime = SIMTINY
RandomVariate = rtime
End Function

Function UniformInt(lolimit As Integer, hilimit As Integer) As Integer
UniformInt = CInt(Rnd * (hilimit - lolimit) + lolimit)
End Function
Sub RandomInit(seed As Single)
Rnd (-seed)
Randomize (seed)
ProbDistGlobal.seed = seed
End Sub
Function GetRandomTime(meantime As Double, cv As Double, newseed As Single) As Double
Dim stddev As Double
Dim low As Single
Dim high As Single
stddev = cv * meantime
low = meantime - 3 * stddev
If low < 0 Then low = 0
high = low + 6 * stddev
If high > low + SIMTINY Then
    ProbDistGlobal.DistType = DISTBETA
    ProbDistGlobal.param1 = low
    ProbDistGlobal.param2 = high
    ProbDistGlobal.param3 = CSng(meantime)
    ProbDistGlobal.seed = newseed
    GetRandomTime = RandomVariate(ProbDistGlobal)
Else
    GetRandomTime = meantime
End If
newseed = ProbDistGlobal.seed
End Function

