Option Explicit

Public Sub RotateLinesAndCreateCircle()
    On Error GoTo ErrorHandler
    
    Dim acadApp As AcadApplication
    Dim acadDoc As AcadDocument
    Dim lineA As Object ' AcadLine
    Dim lineB As Object ' AcadLine
    Dim objEntity As Object ' AcadEntity
    Dim lineR1 As Object ' AcadLine
    Dim lineR2 As Object ' AcadLine
    Dim commonPoint As Variant
    Dim foundCommonPoint As Boolean
    Dim tolerance As Double
    Dim layer0 As AcadLayer
    Dim lineA1 As Object ' AcadLine
    Dim lineA2 As Object ' AcadLine
    Dim lineB1 As Object ' AcadLine
    Dim lineB2 As Object ' AcadLine
    Dim angle90 As Double
    Dim angleMinus90 As Double
    Dim objCircle As Object ' AcadCircle
    Dim radius As Double
    
    ' Get the active document
    Set acadApp = ThisDrawing.Application
    Set acadDoc = ThisDrawing
    
    ' Set tolerance for point comparison (small value for coordinate matching)
    tolerance = 0.0001
    
    ' Prompt user to select first line (Line A)
    On Error Resume Next
    acadDoc.Utility.GetEntity objEntity, commonPoint, vbCrLf & "Select Line A: "
    If Err.Number <> 0 Then
        MsgBox "Selection cancelled.", vbInformation
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Check if selected entity is a line
    If objEntity.ObjectName <> "AcDbLine" Then
        MsgBox "Selected entity is not a line. Please select a line.", vbExclamation
        Exit Sub
    End If
    Set lineA = objEntity
    
    ' Prompt user to select second line (Line B)
    On Error Resume Next
    acadDoc.Utility.GetEntity objEntity, commonPoint, vbCrLf & "Select Line B: "
    If Err.Number <> 0 Then
        MsgBox "Selection cancelled.", vbInformation
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Check if selected entity is a line
    If objEntity.ObjectName <> "AcDbLine" Then
        MsgBox "Selected entity is not a line. Please select a line.", vbExclamation
        Exit Sub
    End If
    Set lineB = objEntity
    
    ' Find common endpoint (Point C)
    foundCommonPoint = False
    
    ' Check if Line A's start point matches Line B's start point
    If PointsAreEqual(lineA.StartPoint, lineB.StartPoint, tolerance) Then
        commonPoint = lineA.StartPoint
        foundCommonPoint = True
    ' Check if Line A's start point matches Line B's end point
    ElseIf PointsAreEqual(lineA.StartPoint, lineB.EndPoint, tolerance) Then
        commonPoint = lineA.StartPoint
        foundCommonPoint = True
    ' Check if Line A's end point matches Line B's start point
    ElseIf PointsAreEqual(lineA.EndPoint, lineB.StartPoint, tolerance) Then
        commonPoint = lineA.EndPoint
        foundCommonPoint = True
    ' Check if Line A's end point matches Line B's end point
    ElseIf PointsAreEqual(lineA.EndPoint, lineB.EndPoint, tolerance) Then
        commonPoint = lineA.EndPoint
        foundCommonPoint = True
    End If
    
    ' If no common point found, exit
    If Not foundCommonPoint Then
        MsgBox "The selected lines do not share a common endpoint.", vbExclamation
        Exit Sub
    End If
    
    ' Ensure Layer 0 exists (it should always exist)
    On Error Resume Next
    Set layer0 = acadDoc.Layers.Item("0")
    On Error GoTo ErrorHandler
    
    ' Create the four rotated lines
    
    ' Define rotation angles (in radians)
    angle90 = 90 * 3.14159265358979 / 180      ' 90 degrees
    angleMinus90 = -90 * 3.14159265358979 / 180 ' -90 degrees
    
    ' Copy and rotate Line A by +90 degrees (Line A1)
    Set lineA1 = lineA.Copy
    lineA1.Layer = "0"
    Call RotateLineAroundPoint(lineA1, commonPoint, angle90)
    Call EnsureStartsAtPoint(lineA1, commonPoint, tolerance)
    
    ' Copy and rotate Line A by -90 degrees (Line A2)
    Set lineA2 = lineA.Copy
    lineA2.Layer = "0"
    Call RotateLineAroundPoint(lineA2, commonPoint, angleMinus90)
    Call EnsureStartsAtPoint(lineA2, commonPoint, tolerance)
    
    ' Copy and rotate Line B by +90 degrees (Line B1)
    Set lineB1 = lineB.Copy
    lineB1.Layer = "0"
    Call RotateLineAroundPoint(lineB1, commonPoint, angle90)
    Call EnsureStartsAtPoint(lineB1, commonPoint, tolerance)
    
    ' Copy and rotate Line B by -90 degrees (Line B2)
    Set lineB2 = lineB.Copy
    lineB2.Layer = "0"
    Call RotateLineAroundPoint(lineB2, commonPoint, angleMinus90)
    Call EnsureStartsAtPoint(lineB2, commonPoint, tolerance)

    ' Add center lines in red (pointwise midpoint between corresponding ends)
    Dim midStart1 As Variant
    Dim midEnd1 As Variant
    Dim midStart2 As Variant
    Dim midEnd2 As Variant

    midStart1 = MidPoint(lineA1.StartPoint, lineB1.StartPoint)
    midEnd1 = MidPoint(lineA1.EndPoint, lineB1.EndPoint)
    midStart2 = MidPoint(lineA2.StartPoint, lineB2.StartPoint)
    midEnd2 = MidPoint(lineA2.EndPoint, lineB2.EndPoint)

    Set lineR1 = acadDoc.ModelSpace.AddLine(midStart1, midEnd1)
    lineR1.Layer = "0"
    lineR1.Color = 1 ' red

    Set lineR2 = acadDoc.ModelSpace.AddLine(midStart2, midEnd2)
    lineR2.Layer = "0"
    lineR2.Color = 1 ' red
    
    ' Create objCircle at Point C with 1.3mm radius
    radius = 1.3 ' 1.3mm radius
    
    Set objCircle = acadDoc.ModelSpace.AddCircle(commonPoint, radius)
    objCircle.Layer = "0"
    
    ' Refresh the display
    acadDoc.Regen acAllViewports
    
   '  MsgBox "Operation completed successfully!" & vbCrLf & "Created 4 rotated lines and 1 objCircle at the common endpoint.", vbInformation
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Helper function to check if two points are equal within tolerance
Private Function PointsAreEqual(point1 As Variant, point2 As Variant, tolerance As Double) As Boolean
    Dim dx As Double, dy As Double, dz As Double
    
    dx = Abs(point1(0) - point2(0))
    dy = Abs(point1(1) - point2(1))
    dz = Abs(point1(2) - point2(2))
    
    If dx < tolerance And dy < tolerance And dz < tolerance Then
        PointsAreEqual = True
    Else
        PointsAreEqual = False
    End If
End Function

' Helper subroutine to rotate a line around a point
Private Sub RotateLineAroundPoint(line As AcadLine, basePoint As Variant, angle As Double)
    Dim newStartPoint As Variant
    Dim newEndPoint As Variant
    
    ' Rotate start point
    newStartPoint = RotatePointAroundBase(line.StartPoint, basePoint, angle)
    
    ' Rotate end point
    newEndPoint = RotatePointAroundBase(line.EndPoint, basePoint, angle)
    
    ' Update line coordinates
    line.StartPoint = newStartPoint
    line.EndPoint = newEndPoint
End Sub

' Helper function to rotate a point around a base point
Private Function RotatePointAroundBase(point As Variant, basePoint As Variant, angle As Double) As Variant
    Dim dx As Double, dy As Double
    Dim newX As Double, newY As Double
    Dim cosAngle As Double, sinAngle As Double
    Dim result(0 To 2) As Double
    
    ' Calculate relative position from base point
    dx = point(0) - basePoint(0)
    dy = point(1) - basePoint(1)
    
    ' Calculate cos and sin of rotation angle
    cosAngle = Cos(angle)
    sinAngle = Sin(angle)
    
    ' Apply rotation matrix
    newX = dx * cosAngle - dy * sinAngle
    newY = dx * sinAngle + dy * cosAngle
    
    ' Translate back to original coordinate system
    result(0) = basePoint(0) + newX
    result(1) = basePoint(1) + newY
    result(2) = point(2) ' Keep Z coordinate unchanged
    
    RotatePointAroundBase = result
End Function

Private Function MidPoint(p1 As Variant, p2 As Variant) As Variant
    Dim m(0 To 2) As Double
    m(0) = (p1(0) + p2(0)) / 2
    m(1) = (p1(1) + p2(1)) / 2
    m(2) = (p1(2) + p2(2)) / 2
    MidPoint = m
End Function

Private Sub EnsureStartsAtPoint(ln As Object, targetPoint As Variant, tol As Double)
    ' Ensure the line's StartPoint is the target; if not, swap start/end
    If Not PointsAreEqual(ln.StartPoint, targetPoint, tol) Then
        Dim tmpStart As Variant
        tmpStart = ln.StartPoint
        ln.StartPoint = ln.EndPoint
        ln.EndPoint = tmpStart
    End If
End Sub
