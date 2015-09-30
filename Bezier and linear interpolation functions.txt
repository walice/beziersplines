Option Explicit

Function Bezier(KnownXs As Range, KnownYs As Range, X As Double, Optional Extrapolate As Integer) As Variant
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
'This function allows you to interpolate Y values by replicating Excel's smoothing algorithm for its
'smooth line scatter plot.
'It creates a third order Bezier curve and interpolates from the relevant spline segment.
'There is an extra option to extrapolate if the X value is outside of the range of the known X values.

'Inspiration: http://blog.splitwise.com/2012/01/31/mystery-solved-the-secret-of-excel-curved-line-interpolation

'ALICE LEPISSIER, Center for Global Development, alepissier@cgdev.org
'October 2014
'This code is free and open-source. You are free to run the code for any purpose, modify it and redistribute
'it. This code is provided in the hope that it will be useful, but without any warranty; without even the
'implied warranty of merchantability or fitness for a particular purpose.
'Feedback is most welcome. Please preserve the comments in the code if you are redistributing it.
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////


'///////////////////////////////////////////////////////
'ERROR TRAPPING
'///////////////////////////////////////////////////////

'Check if the X and Y vectors are the same length, and if there are enough data points for a Bezier curve.
Dim nR As Integer
nR = KnownXs.Rows.Count

If nR <> KnownYs.Rows.Count Then
    GoTo NotSameRange
ElseIf nR < 4 Then
    GoTo NotBezier
End If


'Check if X values are monotonically increasing.
Dim j As Integer
Dim bMono As Boolean

For j = 1 To nR - 1
    If KnownXs(j, 1) <= KnownXs(j + 1) Then
        bMono = True
    Else: bMono = False
    End If
Next j

If bMono = False Then
    GoTo NotMonotonic
End If


'Return Y value if X value already exists.
For j = 1 To nR
    If X = KnownXs(j) Then
        Bezier = KnownYs(j)
        Exit Function
    End If
Next


'///////////////////////////////////////////////////////
'OPTIONAL ARGUMENT TO EXTRAPOLATE
'///////////////////////////////////////////////////////

Dim bUnique As Boolean

If Extrapolate <> 1 And (X > KnownXs(nR) Or X < KnownXs(1)) Then
    GoTo OutsideRange
End If

If Extrapolate = 1 Then
    If X > KnownXs(nR) Then
    'Extrapolate forward
        For j = 1 To nR - 1
            If Not (KnownXs(nR - 1) < KnownXs(nR)) Then
                bUnique = False
            Else: bUnique = True
            End If
        Next
        
    If bUnique = False Then
        GoTo NotUniquelyValued
    End If
    
    Bezier = KnownYs(nR - 1) + _
    (KnownYs(nR) - KnownYs(nR - 1)) / _
    (KnownXs(nR) - KnownXs(nR - 1)) * _
    (X - KnownXs(nR - 1))
    Exit Function
    
    ElseIf X < KnownXs(1) Then
    'Extrapolate backward
        For j = 1 To nR - 1
            If Not (KnownXs(1) < KnownXs(2)) Then
                bUnique = False
            Else: bUnique = True
            End If
        Next
        
    If bUnique = False Then
        GoTo NotUniquelyValued
    End If
    
    Bezier = KnownYs(1) + _
    (KnownYs(2) - KnownYs(1)) / _
    (KnownXs(2) - KnownXs(1)) * _
    (X - KnownXs(1))
    Exit Function
    
    End If
End If


'///////////////////////////////////////////////////////
'CONSTRUCTING THE BEZIER CURVES
'///////////////////////////////////////////////////////

'First find which segment the data point is in.
Dim S, Segment As Integer

S = Application.Match(X, KnownXs, 1)

If S >= KnownYs.Rows.Count - 1 Then
    Segment = 3
ElseIf S < 2 Then
    Segment = 1
Else
    Segment = 2
End If
'Debug.Print S, Segment


'Assign the value to interpolate to the relevant control points.
Dim Ax, Bx, Cx, Dx, Ay, By, Cy, Dy As Variant

Select Case Segment
    Case 1
    'This is the first segment
    Ax = KnownXs(S, 1)
    Bx = KnownXs(S + 1, 1)
    Cx = KnownXs(S + 2, 1)
    Dx = KnownXs(S + 3, 1)
    Ay = KnownYs(S, 1)
    By = KnownYs(S + 1, 1)
    Cy = KnownYs(S + 2, 1)
    Dy = KnownYs(S + 3, 1)

    Case 2
    'This is a middle segment
    Ax = KnownXs(S - 1, 1)
    Bx = KnownXs(S, 1)
    Cx = KnownXs(S + 1, 1)
    Dx = KnownXs(S + 2, 1)
    Ay = KnownYs(S - 1, 1)
    By = KnownYs(S, 1)
    Cy = KnownYs(S + 1, 1)
    Dy = KnownYs(S + 2, 1)

    Case 3
    'This is the last segment
    Ax = KnownXs(S - 2, 1)
    Bx = KnownXs(S - 1, 1)
    Cx = KnownXs(S, 1)
    Dx = KnownXs(S + 1, 1)
    Ay = KnownYs(S - 2, 1)
    By = KnownYs(S - 1, 1)
    Cy = KnownYs(S, 1)
    Dy = KnownYs(S + 1, 1)
End Select
'Debug.Print Ax; Bx; Cx; Dx; Ay; By; Cy; Dy


'Create the distance vectors between the control points.
Dim Zero1, One2, Two3, Zero2, One3 As Variant
Zero1 = ((Ax - Bx) ^ 2 + (Ay - By) ^ 2) ^ 0.5
One2 = ((Bx - Cx) ^ 2 + (By - Cy) ^ 2) ^ 0.5
Two3 = ((Cx - Dx) ^ 2 + (Cy - Dy) ^ 2) ^ 0.5
Zero2 = ((Ax - Cx) ^ 2 + (Ay - Cy) ^ 2) ^ 0.5
One3 = ((Bx - Dx) ^ 2 + (By - Dy) ^ 2) ^ 0.5
'Debug.Print Zero1, One2, Two3, Zero2, One3


'Then compute the control points.
Dim P1ABx, P2ABx, P1BCx, P2BCx, P1CDx, P2CDx, P1ABy, P2ABy, P1BCy, P2BCy, P1CDy, P2CDy As Variant

P1ABx = Ax + (Bx - Ax) * 1 / 6
P2ABx = Bx + (Ax - Cx) * 1 / 6
P1ABy = Ay + (By - Ay) * 1 / 6
P2ABy = By + (Ay - Cy) * 1 / 6
P1CDx = Cx + (Dx - Bx) * 1 / 6
P2CDx = Dx + (Cx - Dx) * 1 / 6
P1CDy = Cy + (Dy - By) * 1 / 6
P2CDy = Dy + (Cy - Dy) * 1 / 6


'Adjust the distance between the control points.
If (Zero2 / 6 < One2 / 2) And (One3 / 6 < One2 / 2) Then
    P1BCx = Bx + (Cx - Ax) * 1 / 6
    P2BCx = Cx + (Bx - Dx) * 1 / 6
    P1BCy = By + (Cy - Ay) * 1 / 6
    P2BCy = Cy + (By - Dy) * 1 / 6
ElseIf (Zero2 / 6 >= One2 / 2) And (One3 / 6 >= One2 / 2) Then
    P1BCx = Bx + (Cx - Ax) * One2 / 2 / Zero2
    P2BCx = Cx + (Bx - Dx) * One2 / 2 / One3
    P1BCy = By + (Cy - Ay) * One2 / 2 / Zero2
    P2BCy = Cy + (By - Dy) * One2 / 2 / One3
ElseIf (Zero2 / 6 >= One2 / 2) Then
    P1BCx = Bx + (Cx - Ax) * One2 / 2 / Zero2
    P2BCx = Cx + (Bx - Dx) * One2 / 2 / One3 * (One3 / Zero2)
    P1BCy = By + (Cy - Ay) * One2 / 2 / Zero2
    P2BCy = Cy + (By - Dy) * One2 / 2 / One3 * (One3 / Zero2)
Else
    P1BCx = Bx + (Cx - Ax) * One2 / 2 / Zero2 * (One2 / One3)
    P2BCx = Cx + (Bx - Dx) * One2 / 2 / One3
    P1BCy = By + (Cy - Ay) * One2 / 2 / Zero2 * (One2 / One3)
    P2BCy = Cy + (By - Dy) * One2 / 2 / One3
End If
'Debug.Print P1ABx; P2ABx; P1BCx; P2BCx; P1CDx; P2CDx
'Debug.Print P1ABy; P2ABy; P1BCy; P2BCy; P1CDy; P2CDy


'Declare an array with the parameter t.
Dim t
t = Array(0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1)


'Loop through t and compute the F'x(t) and G'y(t) parametric curves by adding to the array.
Dim n As Long
Dim ABFx(), ABGy(), BCFx(), BCGy(), CDFx(), CDGy() As Variant
Dim bDimmed As Boolean
Dim bFound As Boolean
Dim P As Integer

bDimmed = False
bFound = False

For n = LBound(t) To UBound(t)

    If bDimmed = True Then
    'The F'x(t) and G'y(t) arrays have been created and we add to the last element
        ReDim Preserve ABFx(0 To UBound(ABFx) + 1) As Variant
        ReDim Preserve ABGy(0 To UBound(ABGy) + 1) As Variant
        ReDim Preserve BCFx(0 To UBound(BCFx) + 1) As Variant
        ReDim Preserve BCGy(0 To UBound(BCGy) + 1) As Variant
        ReDim Preserve CDFx(0 To UBound(CDFx) + 1) As Variant
        ReDim Preserve CDGy(0 To UBound(CDGy) + 1) As Variant
    Else
    'We dimension the arrays and flag them as such
        ReDim ABFx(0 To 0) As Variant
        ReDim ABGy(0 To 0) As Variant
        ReDim BCFx(0 To 0) As Variant
        ReDim BCGy(0 To 0) As Variant
        ReDim CDFx(0 To 0) As Variant
        ReDim CDGy(0 To 0) As Variant
        bDimmed = True
    End If


    'Construct the parametric Bezier curves F'x(t) and G'y(t) with the Bernstein polynomials.
        'These are for the first segment.
    ABFx(UBound(ABFx)) = (Ax * (1 - t(n)) ^ 3 + P1ABx * 3 * t(n) * (1 - t(n)) ^ 2 + P2ABx * 3 * t(n) ^ 2 * (1 - t(n)) + Bx * t(n) ^ 3)
    ABGy(UBound(ABGy)) = (Ay * (1 - t(n)) ^ 3 + P1ABy * 3 * t(n) * (1 - t(n)) ^ 2 + P2ABy * 3 * t(n) ^ 2 * (1 - t(n)) + By * t(n) ^ 3)
        
        'These are for middle segments.
    BCFx(UBound(BCFx)) = (Bx * (1 - t(n)) ^ 3 + P1BCx * 3 * t(n) * (1 - t(n)) ^ 2 + P2BCx * 3 * t(n) ^ 2 * (1 - t(n)) + Cx * t(n) ^ 3)
    BCGy(UBound(BCGy)) = (By * (1 - t(n)) ^ 3 + P1BCy * 3 * t(n) * (1 - t(n)) ^ 2 + P2BCy * 3 * t(n) ^ 2 * (1 - t(n)) + Cy * t(n) ^ 3)
    
        'These are for the last segment.
    CDFx(UBound(CDFx)) = (Cx * (1 - t(n)) ^ 3 + P1CDx * 3 * t(n) * (1 - t(n)) ^ 2 + P2CDx * 3 * t(n) ^ 2 * (1 - t(n)) + Dx * t(n) ^ 3)
    CDGy(UBound(CDGy)) = (Cy * (1 - t(n)) ^ 3 + P1CDy * 3 * t(n) * (1 - t(n)) ^ 2 + P2CDy * 3 * t(n) ^ 2 * (1 - t(n)) + Dy * t(n) ^ 3)
    'Debug.Print ABFx(n); ABGy(n)
    'Debug.Print BCFx(n); BCGy(n)
    'Debug.Print CDFx(n); CDGy(n)


    'Find the closest points on the Bezier curve to interpolate from.
    If bFound = False Then
        Select Case Segment
            Case 1
                If ABFx(n) > X Then
                bFound = True
                P = n
                End If
            Case 2
                If BCFx(n) > X Then
                bFound = True
                P = n
                End If
            Case 3
                If CDFx(n) > X Then
                bFound = True
                P = n
                End If
        End Select
    End If

Next n
'Debug.Print P;


'///////////////////////////////////////////////////////
'INTERPOLATION
'///////////////////////////////////////////////////////
Dim lin As Variant

'We now linearly interpolate between the points on the Bezier curves.
Select Case Segment
    Case 1
    'This is the first segment.
    lin = ABGy(P - 1) + _
    (ABGy(P) - ABGy(P - 1)) / _
    (ABFx(P) - ABFx(P - 1)) * _
    (X - ABFx(P - 1))

    Case 2
    'This is a middle segment.
    lin = BCGy(P - 1) + _
    (BCGy(P) - BCGy(P - 1)) / _
    (BCFx(P) - BCFx(P - 1)) * _
    (X - BCFx(P - 1))
    
    Case 3
    'This is the last segment.
    lin = CDGy(P - 1) + _
    (CDGy(P) - CDGy(P - 1)) / _
    (CDFx(P) - CDFx(P - 1)) * _
    (X - CDFx(P - 1))
End Select

'This is the result.
Bezier = lin
Exit Function


'///////////////////////////////////////////////////////
'ERROR HANDLERS
'///////////////////////////////////////////////////////
NotSameRange:
MsgBox "The number of X values isn't the same as the number of Y values.", , "Warning"
Bezier = CVErr(xlErrRef)
Exit Function

NotBezier:
MsgBox "You need at least 4 data points for Bézier interpolation." _
& Chr(13) & "With less than 3 data points, you can only do linear interpolation." _
& Chr(13) & "Try the Linerp() function.", , "Warning"
Bezier = CVErr(xlErrRef)
Exit Function

NotMonotonic:
MsgBox "The X values need to be monotonically increasing." _
& Chr(13) & "Either sort your X values or interpolate on the Y axis.", , "Error"
Bezier = CVErr(xlErrValue)
Exit Function

NotUniquelyValued:
MsgBox "The endpoint X values need to be uniquely valued for the extrapolation to work.", , "Error"
Bezier = CVErr(xlErrValue)
Exit Function

OutsideRange:
MsgBox "The X value to interpolate is outside the range of known X values." _
& Chr(13) & "Type 1 to include the optional argument to extrapolate backward and forward.", , "Warning"
Bezier = CVErr(xlErrName)
Exit Function


End Function


Function Linerp(KnownXs As Range, KnownYs As Range, X As Double)
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
'This function does piecewise linear interpolation for X and Y data in columns or rows (any order).
'If the X to interpolate is outside the range of known Xs, then the function
'extrapolates backward and forward.
'The function can deal with increasing and decreasing X data, but the data needs to be monotonic.

'ALICE LEPISSIER, Center for Global Development, alepissier@cgdev.org
'October 2014
'This code is free and open-source. You are free to run the code for any purpose, modify it and redistribute
'it. This code is provided in the hope that it will be useful, but without any warranty; without even the
'implied warranty of merchantability or fitness for a particular purpose.
'Feedback is most welcome. Please preserve the comments in the code if you are redistributing it.
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////


'///////////////////////////////////////////////////////
'DIMENSION THE VARIABLES
'///////////////////////////////////////////////////////
Dim Row, Col As Integer
Dim R, C As Integer
Dim nR, nC As Integer
Dim bUnique As Boolean
Dim j As Integer
Dim bMonoInc, bMonoDec, bIncreasing As Integer

bMonoInc = 0
bMonoDec = 0

Row = KnownXs.Rows.Count + KnownYs.Rows.Count
Col = KnownXs.Columns.Count + KnownYs.Columns.Count

nR = KnownXs.Rows.Count
nC = KnownXs.Columns.Count


'///////////////////////////////////////////////////////
'ERROR TRAPPING
'///////////////////////////////////////////////////////
'Check if there are enough data points to interpolate.
If Row < 2 And Col < 2 Then Exit Function

'Check if the X and Y vectors are the same length.
If nC <> KnownYs.Columns.Count Or nR <> KnownYs.Rows.Count Then
    GoTo NotSameRange
End If

If nR = 2 Then
    'Data is in rows
    GoTo ROW_DATA
ElseIf nC = 1 Then
    'Data is in columns
    GoTo COLUMN_DATA
End If


'///////////////////////////////////////////////////////
ROW_DATA:
'///////////////////////////////////////////////////////
'Check if X values are monotonically increasing or decreasing.
For j = 1 To nC - 1
    If KnownXs(, j) <= KnownXs(, j + 1) Then
        bMonoInc = bMonoInc + 1
    End If
    If KnownXs(, j) >= KnownXs(, j + 1) Then
        bMonoDec = bMonoDec + 1
    End If
Next

If bMonoInc < nC - 1 And bMonoDec < nC - 1 Then
    GoTo NotMonotonic
End If

'Check for strict monotonicity.
For j = 1 To nC - 1
    If KnownXs(, j) < KnownXs(, j + 1) Then
        bIncreasing = 1
    ElseIf KnownXs(, j) > KnownXs(, j + 1) Then
        bIncreasing = 0
    End If
Next j

'Extrapolate forward with strictly increasing data.
If X > KnownXs(, nC) And bIncreasing = 1 Then
    If KnownXs(, nC - 1) <> KnownXs(, nC) Then
        Linerp = KnownYs(, nC - 1) + (KnownYs(, nC) - KnownYs(, nC - 1)) / (KnownXs(, nC) - KnownXs(, nC - 1)) * (X - KnownXs(, nC - 1))
        Exit Function
    Else
        GoTo NotUniquelyValued
    End If
End If

'Extrapolate backward with strictly increasing data.
If X < KnownXs(, 1) And bIncreasing = 1 Then
    If KnownXs(, 1) <> KnownXs(, 2) Then
        Linerp = KnownYs(, 1) + (KnownYs(, 2) - KnownYs(, 1)) / (KnownXs(, 2) - KnownXs(, 1)) * (X - KnownXs(, 1))
        Exit Function
    Else
        GoTo NotUniquelyValued
    End If
End If

'Extrapolate forward with strictly decreasing data.
If X > KnownXs(, 1) And bIncreasing = 0 Then
    If KnownXs(, 1) <> KnownXs(, 2) Then
        Linerp = KnownYs(, 1) + (KnownYs(, 2) - KnownYs(, 1)) / (KnownXs(, 2) - KnownXs(, 1)) * (X - KnownXs(, 1))
        Exit Function
    Else
        GoTo NotUniquelyValued
    End If
End If

'Extrapolate backward with strictly decreasing data.
If X < KnownXs(, nC) And bIncreasing = 0 Then
    If KnownXs(, nC - 1) <> KnownXs(, nC) Then
        Linerp = KnownYs(, nC - 1) + (KnownYs(, nC) - KnownYs(, nC - 1)) / (KnownXs(, nC) - KnownXs(, nC - 1)) * (X - KnownXs(, nC - 1))
        Exit Function
    Else
        GoTo NotUniquelyValued
    End If
End If

'Return Y value if X value already exists.
For C = 1 To nC
    If X = KnownXs(, C) Then
        Linerp = KnownYs(, C)
        Exit Function
    End If

'Piecewise linear interpolation.
    If (bIncreasing = 1 And X < KnownXs(, C)) Or (bIncreasing = 0 And X > KnownXs(, C)) Then
        Linerp = KnownYs(, C - 1) + (KnownYs(, C) - KnownYs(, C - 1)) / (KnownXs(, C) - KnownXs(, C - 1)) * (X - KnownXs(, C - 1))
        Exit Function
    End If
Next

Exit Function


'///////////////////////////////////////////////////////
COLUMN_DATA:
'///////////////////////////////////////////////////////
'Check if X values are monotonically increasing or decreasing.
For j = 1 To nR - 1
    If KnownXs(j) <= KnownXs(j + 1) Then
        bMonoInc = bMonoInc + 1
    End If
    If KnownXs(j) >= KnownXs(j + 1) Then
        bMonoDec = bMonoDec + 1
    End If
Next

If bMonoInc < nR - 1 And bMonoDec < nR - 1 Then
    GoTo NotMonotonic
End If

'Check for strict monotonicity.
For j = 1 To nR - 1
    If KnownXs(j) < KnownXs(j + 1) Then
        bIncreasing = 1
    ElseIf KnownXs(j) > KnownXs(j + 1) Then
        bIncreasing = 0
    End If
Next j

'Extrapolate forward with strictly increasing data.
If X > KnownXs(nR) And bIncreasing = 1 Then
    If KnownXs(nR - 1) <> KnownXs(nR) Then
        Linerp = KnownYs(nR - 1) + (KnownYs(nR) - KnownYs(nR - 1)) / (KnownXs(nR) - KnownXs(nR - 1)) * (X - KnownXs(nR - 1))
        Exit Function
    Else
        GoTo NotUniquelyValued
    End If
End If

'Extrapolate backward with strictly increasing data.
If X < KnownXs(1) And bIncreasing = 1 Then
    If KnownXs(1) <> KnownXs(2) Then
        Linerp = KnownYs(1) + (KnownYs(2) - KnownYs(1)) / (KnownXs(2) - KnownXs(1)) * (X - KnownXs(1))
        Exit Function
    Else
        GoTo NotUniquelyValued
    End If
End If

'Extrapolate forward with strictly decreasing data.
If X > KnownXs(1) And bIncreasing = 0 Then
    If KnownXs(1) <> KnownXs(2) Then
        Linerp = KnownYs(1) + (KnownYs(2) - KnownYs(1)) / (KnownXs(2) - KnownXs(1)) * (X - KnownXs(1))
        Exit Function
    Else
        GoTo NotUniquelyValued
    End If
End If

'Extrapolate backward with strictly decreasing data.
If X < KnownXs(nR) And bIncreasing = 0 Then
    If KnownXs(nR - 1) <> KnownXs(nR) Then
        Linerp = KnownYs(nR - 1) + (KnownYs(nR) - KnownYs(nR - 1)) / (KnownXs(nR) - KnownXs(nR - 1)) * (X - KnownXs(nR - 1))
        Exit Function
    Else
        GoTo NotUniquelyValued
    End If
End If

'Return Y value if X value already exists.
For R = 1 To nR
    If X = KnownXs(R) Then
        Linerp = KnownYs(R)
        Exit Function
    End If

'Piecewise linear interpolation.
    If (bIncreasing = 1 And X < KnownXs(R)) Or (bIncreasing = 0 And X > KnownXs(R)) Then
        Linerp = KnownYs(R - 1) + (KnownYs(R) - KnownYs(R - 1)) / (KnownXs(R) - KnownXs(R - 1)) * (X - KnownXs(R - 1))
        Exit Function
    End If
Next

Exit Function


'///////////////////////////////////////////////////////
'ERROR HANDLERS
'///////////////////////////////////////////////////////
NotSameRange:
MsgBox "The number of X values isn't the same as the number of Y values.", , "Warning"
Linerp = CVErr(xlErrRef)
Exit Function

NotMonotonic:
MsgBox "Your X values are not monotonic." _
& Chr(13) & "Either sort your X values or interpolate on the Y axis.", , "Error"
Linerp = CVErr(xlErrValue)
Exit Function

NotUniquelyValued:
MsgBox "The endpoint X values need to be uniquely valued for the extrapolation to work.", , "Error"
Linerp = CVErr(xlErrValue)
Exit Function


End Function