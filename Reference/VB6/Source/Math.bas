Attribute VB_Name = "modMath"
Option Explicit

Public Const PI = 3.14159265358979
Public Const PI2 = 6.28318530717958
Public Const PIDiv2 = 1.5707963267949
Public Const E = 2.71828182845905
Public Const Infinity = 1073741824
Public Const Sqr2 = 1.4142135623731
Public Const GoldenSection = 0.618033988749895
Public Const GoldenRatio = 0.618033988749895
Public Const Ln10 = 2.30258509299405
Public Const Ln2 = 0.693147180559945
Public Const DegreeSign = "°"
Public Const SquareSign = "²"

Public Const ToRadians = PI / 180
Public Const ToDegrees = 180 / PI
Public Const ToRad = PI / 180
Public Const ToDeg = 180 / PI
Public Const Rad = PI / 180
Public Const Deg = 180 / PI
Public Const Epsilon As Double = 0.01

Public Function Distance(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
Distance = Sqr((X2 - X1) * (X2 - X1) + (Y2 - Y1) * (Y2 - Y1))
End Function

Public Function NormalToBezier(ByVal X As Double, ByVal Y As Double, ObjectPoints() As OnePoint) As Double
Const Iter As Boolean = True
Dim BezStep As Double
Dim DX As Double, DY As Double
Dim BS3 As Double, BS2 As Double
Dim MaxDist As Double, MaxT As Double
Dim Resolution As Long, T As Long

Dim dXA As Double, dXB As Double, dXC As Double, dXD As Double
Dim dYA As Double, dYB As Double, dYC As Double, dYD As Double
    
MaxDist = Infinity
Resolution = 100

dXD = ObjectPoints(1).X
dXC = 3 * (ObjectPoints(2).X - ObjectPoints(1).X)
dXB = 3 * (ObjectPoints(3).X - ObjectPoints(2).X) - dXC
dXA = ObjectPoints(4).X - ObjectPoints(1).X - dXB - dXC

dYD = ObjectPoints(1).Y
dYC = 3 * (ObjectPoints(2).Y - ObjectPoints(1).Y)
dYB = 3 * (ObjectPoints(3).Y - ObjectPoints(2).Y) - dYC
dYA = ObjectPoints(4).Y - ObjectPoints(1).Y - dYB - dYC

'T = timeGetTime

If Iter Then
    For BezStep = 0 To 1 Step 1 / Resolution
        BS2 = BezStep * BezStep
        BS3 = BS2 * BezStep
        DX = dXA * BS3 + dXB * BS2 + dXC * BezStep + dXD
        DY = dYA * BS3 + dYB * BS2 + dYC * BezStep + dYD
        If (DX - X) * (DX - X) + (DY - Y) * (DY - Y) < MaxDist Then
            MaxDist = (DX - X) * (DX - X) + (DY - Y) * (DY - Y)
            MaxT = BezStep
        End If
    Next
Else
    'do nothing
End If
NormalToBezier = MaxT
'MsgBox timeGetTime - T
End Function

Public Function DistanceToBezier(ByVal X As Double, ByVal Y As Double, ObjectPoints() As OnePoint) As Double
Dim dXA As Double, dXB As Double, dXC As Double, dXD As Double
Dim dYA As Double, dYB As Double, dYC As Double, dYD As Double
Dim DX As Double, DY As Double, MaxT As Double
    
dXD = ObjectPoints(1).X
dXC = 3 * (ObjectPoints(2).X - ObjectPoints(1).X)
dXB = 3 * (ObjectPoints(3).X - ObjectPoints(2).X) - dXC
dXA = ObjectPoints(4).X - ObjectPoints(1).X - dXB - dXC

dYD = ObjectPoints(1).Y
dYC = 3 * (ObjectPoints(2).Y - ObjectPoints(1).Y)
dYB = 3 * (ObjectPoints(3).Y - ObjectPoints(2).Y) - dYC
dYA = ObjectPoints(4).Y - ObjectPoints(1).Y - dYB - dYC

MaxT = NormalToBezier(X, Y, ObjectPoints)

DX = dXA * MaxT ^ 3 + dXB * MaxT ^ 2 + dXC * MaxT + dXD
DY = dYA * MaxT ^ 3 + dYB * MaxT ^ 2 + dYC * MaxT + dYD
DistanceToBezier = Sqr((DX - X) * (DX - X) + (DY - Y) * (DY - Y))
End Function

Public Function Angle(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal X3 As Double, ByVal Y3 As Double) As Double
Dim A1 As Double, A2 As Double, Ang As Double
A1 = GetAngle(X2, Y2, X1, Y1)
A2 = GetAngle(X2, Y2, X3, Y3)

If A2 < A1 Then
    Ang = A1 - A2
Else
    Ang = A2 - A1
End If
If Ang >= PI Then Ang = 2 * PI - Ang
Angle = Ang
End Function

Public Function OAngle(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal X3 As Double, ByVal Y3 As Double) As Double
Dim A1 As Double, A2 As Double, Ang As Double
A1 = GetAngle(X2, Y2, X1, Y1)
A2 = GetAngle(X2, Y2, X3, Y3)
If A2 < A1 Then A2 = A2 + 2 * PI
OAngle = 2 * PI - A2 + A1
If A1 = A2 Then OAngle = 0
End Function

Public Function GetAngle(ByVal XC As Single, ByVal YC As Single, ByVal X1 As Single, ByVal Y1 As Single) As Double
Dim A As Double
If XC = X1 Then
    A = PI - Sgn(YC - Y1) * PIDiv2
Else
    A = Atn((YC - Y1) / (X1 - XC))
    If X1 < XC Then A = A + PI
    If Y1 > YC And X1 > XC Then A = A + PI2
End If
GetAngle = A
End Function

Public Function SolveSquareEquation(ByVal seA As Double, ByVal seB As Double, ByVal Sec As Double) As TwoNumbers
Dim D As Double
D = seB * seB - 4 * seA * Sec

If D < 0 Or seA = 0 Then
    SolveSquareEquation.n1 = EmptyVar
    SolveSquareEquation.n2 = EmptyVar
    Exit Function
End If

seA = 2 * seA
If D = 0 Then
    SolveSquareEquation.n1 = -seB / seA
    SolveSquareEquation.n2 = EmptyVar
    Exit Function
End If

D = Sqr(D)
SolveSquareEquation.n1 = (-seB - D) / seA
SolveSquareEquation.n2 = (D - seB) / seA

End Function

Public Function Minimum(ParamArray Num() As Variant) As Variant
Dim m As Double, Z As Long
m = Infinity
For Z = 0 To UBound(Num)
    If Num(Z) < m Then m = Num(Z)
Next Z
Minimum = m
End Function

Public Function Maximum(ParamArray Num() As Variant) As Variant
Dim m As Double, Z As Long
m = -Infinity
For Z = 0 To UBound(Num)
    If Num(Z) > m Then m = Num(Z)
Next Z
Maximum = m
End Function

Public Function Infimum(ParamArray Num() As Variant) As Variant
Dim m As Double, Z As Long
m = 1073741824
For Z = 0 To UBound(Num)
    If Num(Z) < m Then m = Num(Z)
Next Z
Infimum = m
End Function

Public Function Supremum(ParamArray Num() As Variant) As Variant
Dim m As Double, Z As Long
m = -1073741824
For Z = 0 To UBound(Num)
    If Num(Z) > m Then m = Num(Z)
Next Z
Supremum = m
End Function

Public Function GetPerpPoint(ByVal X As Double, ByVal Y As Double, ByVal XM As Double, ByVal YM As Double, ByVal XN As Double, ByVal YN As Double) As OnePoint
Dim A As Double, B As Double, C As Double, m As Double
Dim P As OnePoint

If YN = YM Then
    P.X = X
    P.Y = YM
ElseIf XM = XN Then
    P.X = XM
    P.Y = Y
Else
    A = (X - XM) * (X - XM) + (Y - YM) * (Y - YM)
    B = (X - XN) * (X - XN) + (Y - YN) * (Y - YN)
    C = (XM - XN) * (XM - XN) + (YM - YN) * (YM - YN)
    If C <> 0 Then
        m = (A + C - B) / (2 * C)
        P.X = XM + (XN - XM) * m
        P.Y = YM + (YN - YM) * m
    Else
        P.X = XM
        P.Y = YM
    End If
End If

GetPerpPoint = P
End Function

Public Function GetMiddlePoint(Point1 As OnePoint, Point2 As OnePoint) As OnePoint
GetMiddlePoint.X = (Point1.X + Point2.X) / 2
GetMiddlePoint.Y = (Point1.Y + Point2.Y) / 2
End Function

Public Function GetSimmPoint(Point1 As OnePoint, Point2 As OnePoint) As OnePoint
GetSimmPoint.X = Point1.X * 2 - Point2.X
GetSimmPoint.Y = Point1.Y * 2 - Point2.Y
End Function

Public Function GetInvertedPoint(ByVal X As Double, ByVal Y As Double, ByVal XC As Double, ByVal YC As Double, ByVal Rad As Double) As OnePoint
Dim SR As Double
If Rad < 0.001 Or (X = XC And Y = YC) Then
    GetInvertedPoint.X = EmptyVar
    GetInvertedPoint.Y = EmptyVar
    Exit Function
End If

SR = Distance(X, Y, XC, YC)
Rad = (Rad * Rad) / (SR * SR)
GetInvertedPoint.X = XC + (X - XC) * Rad
GetInvertedPoint.Y = YC + (Y - YC) * Rad
End Function

Public Function SolveLinearSystem(ByVal A1 As Double, ByVal B1 As Double, ByVal C1 As Double, ByVal A2 As Double, ByVal B2 As Double, ByVal C2 As Double) As OnePoint
Dim D As Double, DX As Double, DY As Double
D = A1 * B2 - A2 * B1
If D = 0 Then
    SolveLinearSystem.X = EmptyVar
    SolveLinearSystem.Y = EmptyVar
    Exit Function
End If
DX = B1 * C2 - B2 * C1
DY = A2 * C1 - A1 * C2
SolveLinearSystem.X = DX / D
SolveLinearSystem.Y = DY / D
End Function

Public Function GetIntersectionOfLines(Line1 As TwoPoints, Line2 As TwoPoints) As OnePoint
Dim A1 As Double, B1 As Double, C1 As Double, A2 As Double, B2 As Double, C2 As Double

With Line1
    A1 = .P2.Y - .P1.Y
    B1 = .P1.X - .P2.X
    C1 = .P2.X * .P1.Y - .P1.X * .P2.Y
End With
With Line2
    A2 = .P2.Y - .P1.Y
    B2 = .P1.X - .P2.X
    C2 = .P2.X * .P1.Y - .P1.X * .P2.Y
End With

GetIntersectionOfLines = SolveLinearSystem(A1, B1, C1, A2, B2, C2)
End Function

Public Function GetIntersectionOfCircles(Center1 As OnePoint, R1 As Double, Center2 As OnePoint, R2 As Double) As TwoPoints
Dim TN As TwoNumbers
Dim TP As TwoPoints
Dim seA As Double, seB As Double, Sec As Double, cK As Double, CB As Double, tSqr As Double
Dim R3 As Double
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double

X1 = Center1.X
Y1 = Center1.Y
X2 = Center2.X
Y2 = Center2.Y
R3 = Distance(X1, Y1, X2, Y2)

If R3 = 0 Then
    GetIntersectionOfCircles = EmptyTwoPoints
    Exit Function
End If

If Y1 = Y2 Then
    If R1 + R2 > R3 + WorldTransform.Epsilon And R1 + R3 > R2 + WorldTransform.Epsilon And R2 + R3 > R1 + WorldTransform.Epsilon And X1 <> X2 Then
        X3 = (R1 * R1 + R3 * R3 - R2 * R2) / (2 * R3)
        tSqr = Sqr(R1 * R1 - X3 * X3)
        TP.P1.X = X1 + (X2 - X1) * X3 / R3
        TP.P1.Y = Y1 - tSqr
        TP.P2.X = TP.P1.X
        TP.P2.Y = Y1 + tSqr
        If X2 < X1 Then Swap TP.P1.Y, TP.P2.Y
        GetIntersectionOfCircles = TP
    ElseIf (Abs(R1 + R2 - R3) <= WorldTransform.Epsilon) Then
        TP.P1.X = X1 + Sgn(X2 - X1) * R1
        TP.P1.Y = Y1
        TP.P2 = TP.P1 'EmptyVar
        'tP.P2.Y = EmptyVar
        GetIntersectionOfCircles = TP
    ElseIf (Abs(R1 + R3 - R2) <= WorldTransform.Epsilon) Or (Abs(R3 + R2 - R1) <= WorldTransform.Epsilon) Then
        TP.P1.X = X1 + Sgn(X2 - X1) * Sgn(R1 - R2) * R1
        TP.P1.Y = Y1
        TP.P2 = TP.P1 'EmptyVar
        'tP.P2.Y = EmptyVar
        GetIntersectionOfCircles = TP
    Else
        GetIntersectionOfCircles = EmptyTwoPoints
    End If
    Exit Function
End If

If (Abs(R1 + R2 - R3) <= WorldTransform.Epsilon) Then
    R3 = R1 / R3
    TP.P1.X = X1 + (X2 - X1) * R3
    TP.P1.Y = Y1 + (Y2 - Y1) * R3
    TP.P2 = TP.P1
    GetIntersectionOfCircles = TP
    Exit Function
End If

If ((Abs(R1 + R3 - R2) <= WorldTransform.Epsilon)) Or ((Abs(R2 + R3 - R1) <= WorldTransform.Epsilon)) Then
    R3 = R1 / R3 * Sgn(R1 - R2)
    TP.P1.X = X1 + (X2 - X1) * R3
    TP.P1.Y = Y1 + (Y2 - Y1) * R3
    TP.P2 = TP.P1
    GetIntersectionOfCircles = TP
    Exit Function
End If

cK = -(X2 - X1) / (Y2 - Y1)
CB = ((R1 - R2) * (R1 + R2) + (X2 - X1) * (X2 + X1) + (Y2 - Y1) * (Y2 + Y1)) / (2 * (Y2 - Y1))
seA = cK * cK + 1
seB = 2 * (cK * CB - X1 - cK * Y1)
Sec = X1 * X1 + CB * CB - 2 * CB * Y1 + Y1 * Y1 - R1 * R1
TN = SolveSquareEquation(seA, seB, Sec)
TP.P1.X = TN.n1
TP.P1.Y = TN.n1 * cK + CB
TP.P2.X = TN.n2
TP.P2.Y = TN.n2 * cK + CB
If Y2 > Y1 Then SwapTwoPoints TP
GetIntersectionOfCircles = TP
End Function

Public Function GetIntersectionOfCircleAndLine(Center1 As OnePoint, Rad As Double, Line1 As TwoPoints) As TwoPoints
Dim R As TwoPoints
Dim P As OnePoint
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double
Dim H As Double, A As Double, A1 As Double, S As Double

R.P1.X = EmptyVar
R.P1.Y = EmptyVar
R.P2.X = EmptyVar
R.P2.Y = EmptyVar

P = GetPerpPoint(Center1.X, Center1.Y, Line1.P1.X, Line1.P1.Y, Line1.P2.X, Line1.P2.Y)
H = Round(Distance(Center1.X, Center1.Y, P.X, P.Y), 4)
Rad = Round(Rad, 4)

If Abs(H - Rad) <= WorldTransform.Epsilon Then
    R.P1.X = P.X
    R.P1.Y = P.Y
    R.P2 = R.P1
ElseIf H < Rad And H > 0 Then
    S = Sqr((Rad - H) * (Rad + H))
    A = Distance(Line1.P1.X, Line1.P1.Y, P.X, P.Y)
    A1 = Distance(Line1.P2.X, Line1.P2.Y, P.X, P.Y)
    If (A > A1) Or (A > S And A1 > S) Then
        S = S / A
        R.P1.X = P.X - (Line1.P1.X - P.X) * S
        R.P1.Y = P.Y - (Line1.P1.Y - P.Y) * S
    Else
        S = S / A1
        R.P1.X = P.X + (Line1.P2.X - P.X) * S
        R.P1.Y = P.Y + (Line1.P2.Y - P.Y) * S
    End If
    R.P2.X = 2 * P.X - R.P1.X 'P.X - (X1 - P.X) * S
    R.P2.Y = 2 * P.Y - R.P1.Y 'P.Y - (Y1 - P.Y) * S
ElseIf H = 0 Then
    A = Distance(Line1.P1.X, Line1.P1.Y, Line1.P2.X, Line1.P2.Y)
    If A <> 0 Then
        S = Rad / A
        R.P1.X = Center1.X + (Line1.P2.X - Line1.P1.X) * S
        R.P1.Y = Center1.Y + (Line1.P2.Y - Line1.P1.Y) * S
        R.P2.X = 2 * Center1.X - R.P1.X
        R.P2.Y = 2 * Center1.Y - R.P1.Y
    End If
'Else
    'do nothing
End If

GetIntersectionOfCircleAndLine = R
End Function

Public Function GetPerpendicularLine(ByVal X As Double, ByVal Y As Double, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As TwoPoints
Dim R As TwoPoints
R = GetPerpendicularLineAbsolute(X, Y, X1, Y1, X2, Y2)
GetPerpendicularLine = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
End Function

Public Function GetPerpendicularLineAbsolute(ByVal X As Double, ByVal Y As Double, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As TwoPoints
Dim R As TwoPoints
Dim P As OnePoint

R.P1.X = X
R.P1.Y = Y
R.P2.X = X + Y2 - Y1
R.P2.Y = Y - X2 + X1
GetPerpendicularLineAbsolute = R
'P = GetPerpPoint(X, Y, X1 - ParDist, Y1, X2 - ParDist, Y2)
'If Y1 = Y2 Then P = GetPerpPoint(X, Y, X1, Y1 - ParDist, X2, Y2 - ParDist)
'R.P1.X = P.X
'R.P1.Y = P.Y
'
'P = GetPerpPoint(X, Y, X1 + ParDist, Y1, X2 + ParDist, Y2)
'If Y1 = Y2 Then P = GetPerpPoint(X, Y, X1, Y1 + ParDist, X2, Y2 + ParDist)
'R.P2.X = P.X
'R.P2.Y = P.Y
'
'If Y2 < Y1 Then SwapTwoPoints R
'If Y1 = Y2 And X2 > X1 Then SwapTwoPoints R
'GetPerpendicularLineAbsolute = R
End Function

Public Function GetGeneralAnLine(ByVal A As Double, ByVal B As Double, ByVal C As Double) As TwoPoints
Dim R As TwoPoints
R = GetGeneralAnLineAbsolute(A, B, C)
GetGeneralAnLine = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
'R.P1.X = EmptyVar
'R.P1.Y = EmptyVar
'R.P2.X = EmptyVar
'R.P2.Y = EmptyVar
'
'If A <> 0 Then
'    If B <> 0 Then
'        If C <> 0 Then
'            R.P1.X = -C / A
'            R.P1.Y = 0
'            R.P2.X = 0
'            R.P2.Y = -C / B
'        Else
'            R.P1.X = 0
'            R.P1.Y = 0
'            R.P2.X = ParDist
'            R.P2.Y = -A / B * ParDist
'        End If
'    Else
'        R.P1.X = -C / A
'        R.P1.Y = -ParDist
'        R.P2.X = R.P1.X
'        R.P2.Y = ParDist
'    End If
'Else
'    If B <> 0 Then
'        R.P2.X = -ParDist
'        R.P2.Y = -C / B
'        R.P1.X = ParDist
'        R.P1.Y = R.P2.Y
'    End If
'End If
'
End Function

Public Function GetGeneralAnLineAbsolute(ByVal A As Double, ByVal B As Double, ByVal C As Double) As TwoPoints
Dim R As TwoPoints
R.P1.X = EmptyVar
R.P1.Y = EmptyVar
R.P2.X = EmptyVar
R.P2.Y = EmptyVar

If A <> 0 Then
    If B <> 0 Then
        If C <> 0 Then
            R.P1.X = -C / A
            R.P1.Y = 0
            R.P2.X = 0
            R.P2.Y = -C / B
        Else
            R.P1.X = 0
            R.P1.Y = 0
            R.P2.X = 1
            R.P2.Y = -A / B
        End If
    Else
        R.P1.X = -C / A
        R.P1.Y = -1
        R.P2.X = R.P1.X
        R.P2.Y = 1
    End If
Else
    If B <> 0 Then
        R.P2.X = -1
        R.P2.Y = -C / B
        R.P1.X = 1
        R.P1.Y = R.P2.Y
    End If
End If

GetGeneralAnLineAbsolute = R
End Function

Public Function GetCanonicAnLine(ByVal X0 As Double, ByVal Y0 As Double, ByVal A1 As Double, ByVal A2 As Double) As TwoPoints
Dim R As TwoPoints
R = GetCanonicAnLineAbsolute(X0, Y0, A1, A2)

GetCanonicAnLine = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
End Function

Public Function GetCanonicAnLineAbsolute(ByVal X0 As Double, ByVal Y0 As Double, ByVal A1 As Double, ByVal A2 As Double) As TwoPoints
Dim R As TwoPoints
R.P1.X = X0
R.P1.Y = Y0
R.P2.X = X0 + A1
R.P2.Y = Y0 + A2
GetCanonicAnLineAbsolute = R
End Function

Public Function GetNormalPointAnLine(ByVal X0 As Double, ByVal Y0 As Double, ByVal A1 As Double, ByVal A2 As Double) As TwoPoints
Dim R As TwoPoints
R = GetNormalPointAnLineAbsolute(X0, Y0, A1, A2)
GetNormalPointAnLine = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
End Function

Public Function GetNormalPointAnLineAbsolute(ByVal X0 As Double, ByVal Y0 As Double, ByVal A1 As Double, ByVal A2 As Double) As TwoPoints
Dim R As TwoPoints
R.P1.X = X0
R.P1.Y = Y0
R.P2.X = X0 - A2
R.P2.Y = Y0 + A1
GetNormalPointAnLineAbsolute = R
End Function

Public Function GetNormalAnLine(ByVal Ang As Double, ByVal D As Double) As TwoPoints
Dim R As TwoPoints
R = GetNormalAnLineAbsolute(Ang, D)
GetNormalAnLine = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
End Function

Public Function GetNormalAnLineAbsolute(ByVal Ang As Double, ByVal D As Double) As TwoPoints
Dim R As TwoPoints, CosAng As Double, SinAng As Double
CosAng = Cos(Ang)
SinAng = Sin(Ang)
R.P1.X = D * CosAng
R.P1.Y = D * SinAng
R.P2.X = R.P1.X - SinAng
R.P2.Y = R.P1.Y + CosAng
GetNormalAnLineAbsolute = R
End Function

Public Function GetLineGeneralEquation(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As LineGeneralEquation
GetLineGeneralEquation.A = Y2 - Y1
GetLineGeneralEquation.B = X1 - X2
GetLineGeneralEquation.C = X2 * Y1 - X1 * Y2
End Function

Public Function GetCircleRadiusFromEquation(ByVal A As Double, ByVal B As Double, ByVal C As Double) As Double
GetCircleRadiusFromEquation = Sqr((A * A + B * B) / 4 - C)
End Function

Public Function GetCircleCenterFromEquation(ByVal A As Double, ByVal B As Double, Optional ByVal C As Double) As OnePoint
GetCircleCenterFromEquation.X = -A / 2
GetCircleCenterFromEquation.Y = -B / 2
End Function

Public Function PointBelongsToFigure(ByVal X As Double, ByVal Y As Double, ByVal Figure1 As Long, Optional ByVal IncludeHidden As Boolean = True) As Boolean
On Local Error Resume Next
Dim P As OnePoint
Dim R As TwoPoints
Dim tB As Boolean
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, X5 As Double, Y5 As Double
Dim A1 As Double, A2 As Double, A3 As Double, Rad As Double
Dim ValidDistance As Double

If Not IsFigure(Figure1) Then Exit Function
If Not Figures(Figure1).Visible Then Exit Function
If Figures(Figure1).Hide And Not IncludeHidden Then Exit Function

ValidDistance = Figures(Figure1).DrawWidth
ToLogicalLength ValidDistance
ValidDistance = ValidDistance + Sensitivity

Select Case Figures(Figure1).FigureType
    Case dsSegment
        X1 = BasePoint(Figures(Figure1).Points(0)).X
        Y1 = BasePoint(Figures(Figure1).Points(0)).Y
        X2 = BasePoint(Figures(Figure1).Points(1)).X
        Y2 = BasePoint(Figures(Figure1).Points(1)).Y
        PointBelongsToFigure = PointBelongsToSegment(X, Y, X1, Y1, X2, Y2, ValidDistance)
        Exit Function
        
    Case dsRay
        X1 = BasePoint(Figures(Figure1).Points(0)).X
        Y1 = BasePoint(Figures(Figure1).Points(0)).Y
        X2 = BasePoint(Figures(Figure1).Points(1)).X
        Y2 = BasePoint(Figures(Figure1).Points(1)).Y
        P = GetPerpPoint(X, Y, X1, Y1, X2, Y2)
        If Distance(X, Y, P.X, P.Y) <= ValidDistance Then
            tB = (X1 < X2 And P.X >= X1) Or (X2 <= X1 And P.X <= X1)
            tB = tB And ((Y1 < Y2 And P.Y >= Y1) Or (Y2 <= Y1 And P.Y <= Y1))
            If tB Then PointBelongsToFigure = True: Exit Function
        End If
    
    Case dsLine_2Points, dsBisector, dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine, dsAnLineGeneral, dsAnLineCanonic, dsAnLineNormal, dsAnLineNormalPoint
        R = GetLineCoordinates(Figure1)
        P = GetPerpPoint(X, Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        If Distance(X, Y, P.X, P.Y) <= ValidDistance Then
            PointBelongsToFigure = True
            Exit Function
        End If
        
    Case dsCircle_CenterAndCircumPoint
        X1 = BasePoint(Figures(Figure1).Points(0)).X
        Y1 = BasePoint(Figures(Figure1).Points(0)).Y
        X2 = BasePoint(Figures(Figure1).Points(1)).X
        Y2 = BasePoint(Figures(Figure1).Points(1)).Y
        If Abs(Distance(X, Y, X1, Y1) - Distance(X1, Y1, X2, Y2)) <= ValidDistance Then
            PointBelongsToFigure = True
            Exit Function
        End If
        
    Case dsCircle_CenterAndTwoPoints
        X1 = BasePoint(Figures(Figure1).Points(0)).X
        Y1 = BasePoint(Figures(Figure1).Points(0)).Y
        X2 = BasePoint(Figures(Figure1).Points(1)).X
        Y2 = BasePoint(Figures(Figure1).Points(1)).Y
        X3 = BasePoint(Figures(Figure1).Points(2)).X
        Y3 = BasePoint(Figures(Figure1).Points(2)).Y
        If Abs(Distance(X, Y, X3, Y3) - Distance(X1, Y1, X2, Y2)) <= ValidDistance Then
            PointBelongsToFigure = True
            Exit Function
        End If
        
    Case dsAnCircle
        P = GetCircleCenter(Figure1)
        Rad = GetCircleRadius(Figure1)
        If Abs(Distance(X, Y, P.X, P.Y) - Rad) <= ValidDistance Then
            PointBelongsToFigure = True
            Exit Function
        End If
        
    Case dsCircle_ArcCenterAndRadiusAndTwoPoints
        X1 = BasePoint(Figures(Figure1).Points(0)).X
        Y1 = BasePoint(Figures(Figure1).Points(0)).Y
        X2 = BasePoint(Figures(Figure1).Points(1)).X
        Y2 = BasePoint(Figures(Figure1).Points(1)).Y
        X3 = BasePoint(Figures(Figure1).Points(2)).X
        Y3 = BasePoint(Figures(Figure1).Points(2)).Y
        X4 = BasePoint(Figures(Figure1).Points(3)).X
        Y4 = BasePoint(Figures(Figure1).Points(3)).Y
        X5 = BasePoint(Figures(Figure1).Points(4)).X
        Y5 = BasePoint(Figures(Figure1).Points(4)).Y
        If Abs(Distance(X, Y, X3, Y3) - Distance(X1, Y1, X2, Y2)) <= ValidDistance Then
            A1 = Figures(Figure1).AuxInfo(2)
            A2 = Figures(Figure1).AuxInfo(3)
            A3 = GetAngle(X3, Y3, X, Y)
            If A2 < A1 Then
                If (A3 < A1 And A3 > A2) Then PointBelongsToFigure = True: Exit Function
            Else
                If (A3 < A1 Or A3 > A2) Then PointBelongsToFigure = True: Exit Function
            End If
        End If
        
    Case dsMeasureDistance
        R.P1.X = Figures(Figure1).AuxPoints(3).X + Figures(Figure1).AuxPoints(6).X - Figures(Figure1).AuxPoints(5).X \ 2
        R.P1.Y = Figures(Figure1).AuxPoints(3).Y + Figures(Figure1).AuxPoints(6).Y - Figures(Figure1).AuxPoints(5).Y
        R.P2.X = Figures(Figure1).AuxPoints(3).X + Figures(Figure1).AuxPoints(6).X + Figures(Figure1).AuxPoints(5).X \ 2
        R.P2.Y = Figures(Figure1).AuxPoints(3).Y + Figures(Figure1).AuxPoints(6).Y '+ Figures(Figure1).AuxPoints(5).Y \ 2
        ToPhysical X, Y
        A1 = -Figures(Figure1).AuxInfo(2) * ToRadians
        A2 = GetAngle(Figures(Figure1).AuxPoints(3).X + Figures(Figure1).AuxPoints(6).X, Figures(Figure1).AuxPoints(3).Y + Figures(Figure1).AuxPoints(6).Y, X, Y)
        A3 = Distance(Figures(Figure1).AuxPoints(3).X + Figures(Figure1).AuxPoints(6).X, Figures(Figure1).AuxPoints(3).Y + Figures(Figure1).AuxPoints(6).Y, X, Y)
        A2 = A2 + A1
        X = Figures(Figure1).AuxPoints(3).X + Figures(Figure1).AuxPoints(6).X + A3 * Cos(A2)
        Y = Figures(Figure1).AuxPoints(3).Y + Figures(Figure1).AuxPoints(6).Y - A3 * Sin(A2)
        If PointInRectangle(X, Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y) Then
            PointBelongsToFigure = True
            Exit Function
        End If
    
    Case dsMeasureAngle
        If IsPointInMeasureText(Figure1, X, Y) Then
            PointBelongsToFigure = True
            Exit Function
        End If
        
        If Figures(Figure1).DrawStyle > 0 Then
            ToPhysical X, Y
            If Abs(Distance(Figures(Figure1).AuxPoints(2).X, Figures(Figure1).AuxPoints(2).Y, X, Y) - ((((Figures(Figure1).DrawStyle - 1) Mod 3) + 1) * (2 + Figures(Figure1).DrawWidth) + Figures(Figure1).AuxInfo(2))) < setCursorSensitivity Then
                A1 = GetAngle(Figures(Figure1).AuxPoints(2).X, Figures(Figure1).AuxPoints(2).Y, Figures(Figure1).AuxPoints(1).X, Figures(Figure1).AuxPoints(1).Y)
                A2 = GetAngle(Figures(Figure1).AuxPoints(2).X, Figures(Figure1).AuxPoints(2).Y, Figures(Figure1).AuxPoints(3).X, Figures(Figure1).AuxPoints(3).Y)
                A3 = GetAngle(Figures(Figure1).AuxPoints(2).X, Figures(Figure1).AuxPoints(2).Y, X, Y)
                If AngleBetweenAngles(A3, A1, A2) Then
                    PointBelongsToFigure = True
                    Exit Function
                End If
            End If
        End If
        
        'Round(.AuxInfo(1), setAnglePrecision) & "°"
        
'    Case dsDynamicLocus
'        If BasePoint(Figures(Figure1).Points(0)).Locus > 0 Then
'            If Locuses(BasePoint(Figures(Figure1).Points(0)).Locus).Dynamic Then
'                P = GetPerpPointPolyline(X, Y, Locuses(BasePoint(Figures(Figure1).Points(0)).Locus).LocusPoints)
'                If Distance(X, Y, P.X, P.Y) <= ValidDistance Then PointBelongsToFigure = True: Exit Function
'            End If
'        End If
End Select

PointBelongsToFigure = False
End Function

Public Function IsPointInMeasureText(ByVal Figure1 As Long, ByVal X As Double, ByVal Y As Double) As Boolean
Dim R As TwoPoints
R.P1.X = Figures(Figure1).AuxPoints(4).X + Figures(Figure1).AuxPoints(6).X - Figures(Figure1).AuxPoints(5).X \ 2
R.P1.Y = Figures(Figure1).AuxPoints(4).Y + Figures(Figure1).AuxPoints(6).Y - Figures(Figure1).AuxPoints(5).Y
R.P2.X = Figures(Figure1).AuxPoints(4).X + Figures(Figure1).AuxPoints(6).X + Figures(Figure1).AuxPoints(5).X \ 2
R.P2.Y = Figures(Figure1).AuxPoints(4).Y + Figures(Figure1).AuxPoints(6).Y '+ Figures(Figure1).AuxPoints(5).Y \ 2
ToPhysical X, Y

If Figures(Figure1).DrawStyle < 4 Then
    If PointInRectangle(X, Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y) Then
        IsPointInMeasureText = True
        Exit Function
    End If
End If
IsPointInMeasureText = False
End Function

Public Function PointBelongsToSegment(ByVal X As Double, ByVal Y As Double, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal ValidDistance As Double) As Boolean
Dim P As OnePoint
P = GetPerpPoint(X, Y, X1, Y1, X2, Y2)
If Distance(X, Y, P.X, P.Y) <= ValidDistance Then
    If PointInRectangle(P.X, P.Y, X1, Y1, X2, Y2) Then PointBelongsToSegment = True: Exit Function
End If
End Function

Public Function LinkPointToSegment(ByRef tX As Double, ByRef tY As Double, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal Rad As Double)
Dim P As OnePoint, P2 As OnePoint, Radius As Double

P = GetPerpPoint(tX, tY, X1, Y1, X2, Y2)
Radius = Distance(tX, tY, P.X, P.Y)
If Distance(P.X, P.Y, X1, Y1) < Distance(P.X, P.Y, X2, Y2) Then
    P2.X = X1
    P2.Y = Y1
Else
    P2.X = X2
    P2.Y = Y2
End If
If Radius <= Rad Then
    If Not PointInRectangle(P.X, P.Y, X1, Y1, X2, Y2) Then
        tX = tX + P2.X - P.X
        tY = tY + P2.Y - P.Y
    End If
Else
    If Not PointInRectangle(P.X, P.Y, X1, Y1, X2, Y2) Then
        tX = tX + P2.X - P.X
        tY = tY + P2.Y - P.Y
        P = GetPerpPoint(tX, tY, X1, Y1, X2, Y2)
        Radius = Distance(tX, tY, P.X, P.Y)
        tX = P.X + (tX - P.X) / Radius * Rad
        tY = P.Y + (tY - P.Y) / Radius * Rad
    Else
        tX = P.X + (tX - P.X) / Radius * Rad
        tY = P.Y + (tY - P.Y) / Radius * Rad
    End If
End If
End Function

Public Function LinkPointToPoint(ByRef tX As Double, ByRef tY As Double, ByVal X1 As Double, ByVal Y1 As Double, ByVal Rad As Double)
tX = tX - X1
tY = tY - Y1
If Sqr(tX ^ 2 + tY ^ 2) > Rad Then
    NormalizeVector tX, tY
    tX = tX * Rad
    tY = tY * Rad
End If
tX = tX + X1
tY = tY + Y1
End Function

Public Function Transpose(Num As Variant, TransPos As Transposition) As Variant
Dim Z As Long
For Z = 1 To TransPos.Count
    If TransPos.Element1(Z) = Num Then Transpose = TransPos.Element2(Z): Exit Function
Next
End Function

Public Function TransposeInv(Num As Variant, TransPos As Transposition) As Variant
Dim Z As Long
For Z = 1 To TransPos.Count
    If TransPos.Element2(Z) = Num Then TransposeInv = TransPos.Element1(Z): Exit Function
Next
End Function

Public Function TranspositionInvert(TransPos As Transposition)
Dim Z As Long
For Z = LBound(TransPos.Element1) To UBound(TransPos.Element2)
    Swap TransPos.Element1(Z), TransPos.Element2(Z)
Next
End Function

Public Sub TranspositionAdd(Num1 As Variant, Num2 As Variant, TransPos As Transposition)
With TransPos
    .Count = .Count + 1
    ReDim Preserve .Element1(1 To .Count)
    ReDim Preserve .Element2(1 To .Count)
    .Element1(.Count) = Num1
    .Element2(.Count) = Num2
End With
End Sub

Public Sub TranspositionClear(TransPos As Transposition)
With TransPos
    .Count = 0
    ReDim .Element1(1 To 1)
    ReDim .Element2(1 To 1)
End With
End Sub

Public Function Factorial(ByVal Num As Long) As Double
Dim Product As Double, Z As Long
If Num < 1 Or Num > 100 Then Factorial = 1: Exit Function
Product = 1
For Z = 1 To Num: Product = Product * Z: Next
Factorial = Product
End Function

Public Function Combin(ByVal longN As Long, ByVal longK As Long) As Double
If longN < 1 Or longN > 100 Or longK < 0 Or longK > longN Then Exit Function
Combin = Factorial(longN) / (Factorial(longK) * Factorial(longN - longK))
End Function

'Trigonometric functions
Public Function Ctg(ByVal X As Double) As Double
Ctg = Cos(X) / Sin(X)
End Function

Public Function Sec(ByVal X As Double) As Double
Sec = 1 / Cos(X)
End Function

Public Function Cosec(ByVal X As Double) As Double
Cosec = 1 / Sin(X)
End Function

'Inverse trigonometric functions
Public Function Arcsin(ByVal X As Double) As Double
If X < -1 Or X > 1 Then Exit Function
If Abs(X) = 1 Then
    Arcsin = PI / 2 * Sgn(X)
Else
    Arcsin = Atn(X / Sqr(1 - X * X))
End If
End Function

Public Function Arccos(ByVal X As Double) As Double
If X < -1 Or X > 1 Then Exit Function
If X = 0 Then
    Arccos = PI / 2
Else
    Arccos = PI / 2 - Arcsin(X)
End If
End Function

Public Function Arcctg(ByVal X As Double) As Double
Arcctg = PI / 2 - Atn(X)
End Function

Public Function ArcSec(ByVal X As Double) As Double
ArcSec = Atn(X / Sqr(1 - X * X)) + (Sgn(X) - 1) * PI / 2
End Function

Public Function ArcCsc(ByVal X As Double) As Double
ArcCsc = Atn(1 / Sqr(1 - X * X)) + (Sgn(X) - 1) * PI / 2
End Function

'Hyperbolic functions
Public Function SinH(ByVal X As Double) As Double
SinH = (Exp(X) - Exp(-X)) / 2
End Function

Public Function CosH(ByVal X As Double) As Double
CosH = (Exp(X) + Exp(-X)) / 2
End Function

Public Function TanH(ByVal X As Double) As Double
TanH = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
End Function

Public Function CotH(ByVal X As Double) As Double
CotH = Exp(-X) / (Exp(X) - Exp(-X)) * 2 + 1
End Function

Public Function SecH(ByVal X As Double) As Double
SecH = 2 / (Exp(X) + Exp(-X))
End Function

Public Function CscH(ByVal X As Double) As Double
CscH = 2 / (Exp(X) - Exp(-X))
End Function

'Area functions hyperbolic
Public Function ArcsinH(ByVal X As Double) As Double
ArcsinH = Log(X + Sqr(X * X + 1))
End Function

Public Function ArccosH(ByVal X As Double) As Double
If X >= 1 Then ArccosH = Log(X + Sqr(X * X - 1))
End Function

Public Function ArctanH(ByVal X As Double) As Double
If X <> 1 And Sgn(X + 1) = Sgn(1 - X) Then ArctanH = Log((1 + X) / (1 - X)) / 2
End Function

Public Function ArccotH(ByVal X As Double) As Double
If Sgn(X + 1) = Sgn(X - 1) Then ArccotH = Log((X + 1) / (X - 1)) / 2
End Function

Public Function ArcsecH(ByVal X As Double) As Double
If X <> 0 And Abs(X) <= 1 Then ArcsecH = Log((Sqr(1 - X * X) + 1) / X)
End Function

Public Function ArccscH(ByVal X As Double) As Double
If X <> 0 And Sgn(X) * Sqr(X * X + 1) > -1 Then ArccscH = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
End Function
'/Areafunctions hyperbolic

Public Function Random(ByVal LB As Double, UB As Double) As Double
Random = Rnd * (UB - LB) + LB
End Function

Public Sub NormalizeVector(X As Variant, Y As Variant)
Dim Norm As Double
Norm = Sqr(X * X + Y * Y)
X = X / Norm
Y = Y / Norm
End Sub

Public Function GetCircleEquationText(CCent As OnePoint, ByVal Rad As Double, Optional ByVal UseSquareSign As Boolean = False) As String
Dim S As String
Dim SqSign As String
If UseSquareSign Then SqSign = SquareSign Else SqSign = "^2"

If CCent.X <> 0 Then
    S = "(x " & IIf(CCent.X > 0, "-", "+") & " " & Format(Abs(CCent.X), setFormatNumber) & ")" & SqSign
Else
    S = "x" & SqSign
End If
If CCent.Y <> 0 Then
    S = S & " + (y " & IIf(CCent.Y > 0, "-", "+") & " " & Format(Abs(CCent.Y), setFormatNumber) & ")" & SqSign
Else
    S = S & " + y" & SqSign
End If
S = S & " = " & Format(Rad, setFormatNumber)
If Rad <> 1 Then S = S & SqSign
GetCircleEquationText = S
End Function

Public Function GetLineEquationText(R As TwoPoints) As String
Dim Eq As LineGeneralEquation
Eq = GetLineGeneralEquation(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
GetLineEquationText = FormatLineGeneralEquation(Eq)
End Function

Public Function Area(Points() As OnePoint) As Double
Dim S As Double, Z As Long

If UBound(Points) = 2 Then
    Area = Distance(Points(1).X, Points(1).Y, Points(2).X, Points(2).Y) ^ 2 * PI
    Exit Function
End If

For Z = 1 To UBound(Points) - 1
    S = S + (Points(Z + 1).X - Points(Z).X) * (Points(Z + 1).Y + Points(Z).Y) / 2
Next
Z = UBound(Points)
S = S + (Points(1).X - Points(Z).X) * (Points(1).Y + Points(Z).Y) / 2

Area = Abs(S)
End Function

Public Function EmptyTwoPoints() As TwoPoints
EmptyTwoPoints.P1.X = EmptyVar
EmptyTwoPoints.P1.Y = EmptyVar
EmptyTwoPoints.P2.X = EmptyVar
EmptyTwoPoints.P2.Y = EmptyVar
End Function

Public Function EmptyOnePoint() As OnePoint
EmptyOnePoint.X = EmptyVar
EmptyOnePoint.Y = EmptyVar
End Function

Public Function GetPerpPointPolyline(ByVal X As Double, ByVal Y As Double, P() As OnePoint) As OnePoint
Dim Z As Long, LB As Long, UB As Long
Dim R As Double, R0 As Double, R1 As Double, R2 As Double
Dim M0 As Long, M1 As Long, M2 As Long
Dim X0 As OnePoint, X1 As OnePoint

LB = LBound(P)
UB = UBound(P)
R1 = Infinity
For Z = LB To UB
    R = Distance(X, Y, P(Z).X, P(Z).Y)
    If R < R1 Then
        R1 = R
        M1 = Z
    End If
Next

If M1 = 0 Then GetPerpPointPolyline = EmptyOnePoint: Exit Function
If UB = LB Then GetPerpPointPolyline = P(M1): Exit Function

R0 = Infinity
R2 = Infinity

If M1 < UB Then M2 = M1 + 1 Else M2 = 1
X1 = GetPerpPoint(X, Y, P(M1).X, P(M1).Y, P(M2).X, P(M2).Y)
If PointInRectangle(X1.X, X1.Y, P(M1).X, P(M1).Y, P(M2).X, P(M2).Y) And Abs(M1 - M2) < UB - LB Then
    R2 = Distance(X, Y, X1.X, X1.Y)
End If

If M1 > 1 Then M0 = M1 - 1 Else M0 = UB
X0 = GetPerpPoint(X, Y, P(M0).X, P(M0).Y, P(M1).X, P(M1).Y)
If PointInRectangle(X0.X, X0.Y, P(M0).X, P(M0).Y, P(M1).X, P(M1).Y) And Abs(M1 - M0) < UB - LB Then
    R0 = Distance(X, Y, X0.X, X0.Y)
End If

If R < R0 And R < R2 Then
    GetPerpPointPolyline = P(M1)
Else
    If R0 <= R2 Then
        GetPerpPointPolyline = X0
    Else
        GetPerpPointPolyline = X1
    End If
End If

End Function

Public Function GetNearestSegmentNum(ByVal X As Double, ByVal Y As Double, P() As OnePoint) As Long
Dim Z As Long, R As Double, R1 As Double, M1 As Long

R1 = Infinity
For Z = LBound(P) To UBound(P)
    R = Distance(X, Y, P(Z).X, P(Z).Y)
    If R < R1 Then
        R1 = R
        M1 = Z
    End If
Next

GetNearestSegmentNum = M1
End Function

Public Function PointInRectangle(ByVal X As Double, ByVal Y As Double, ByVal XM As Double, ByVal YM As Double, ByVal XN As Double, ByVal YN As Double) As Boolean
Dim T As Double
If XM > XN Then
    T = XM
    XM = XN
    XN = T
End If
If YM > YN Then
    T = YM
    YM = YN
    YN = T
End If
If X >= XM And X <= XN And Y >= YM And Y <= YN Then PointInRectangle = True
End Function

Public Function GetBisector(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal X3 As Double, ByVal Y3 As Double) As OnePoint
Dim S1 As Double, S2 As Double
S1 = Distance(X1, Y1, X2, Y2) * 2
S2 = Distance(X3, Y3, X2, Y2) * 2
If S1 = 0 Or S2 = 0 Then GetBisector = EmptyOnePoint: Exit Function
GetBisector.X = X2 + (X1 - X2) / S1 + (X3 - X2) / S2
GetBisector.Y = Y2 + (Y1 - Y2) / S1 + (Y3 - Y2) / S2
End Function

Public Function AngleBetweenAngles(ByVal A As Double, ByVal A1 As Double, ByVal A2 As Double) As Boolean
If A2 < A1 Then Swap A1, A2
If A2 - A1 >= PI Then
    If A >= A2 Or A <= A1 Then AngleBetweenAngles = True
Else
    If A >= A1 And A <= A2 Then AngleBetweenAngles = True
End If
End Function
