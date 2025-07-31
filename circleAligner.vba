' BricsCAD VBA Macro: Align Larger Circles to Smaller Circle Centers
' Iterates all circles in the active document, finds larger circles whose centers are inside smaller circles, and moves them to the smaller circle's center

Public Sub CircleAligner()
    Dim doc As AcadDocument
    Set doc = ThisDrawing
    
    Dim circles As Collection
    Set circles = New Collection
    
    Dim ent As AcadEntity
    ' Collect all circles in the drawing
    For Each ent In doc.ModelSpace
        If ent.ObjectName = "AcDbCircle" Then
            circles.Add ent
        End If
    Next
    
    Dim i As Integer, j As Integer
    Dim c1 As AcadCircle, c2 As AcadCircle
    
    ' Iterate through all pairs of circles
    For i = 1 To circles.Count
        Set c1 = circles(i)
        For j = 1 To circles.Count
            If i <> j Then
                Set c2 = circles(j)
                ' Check if c2 center is inside c1 and c2 has bigger radius
                If c2.Radius > c1.Radius Then
                    If IsPointInsideCircle(c1, c2.Center) Then
                        ' Move c2 to c1's center
                        c2.Center = c1.Center
                    End If
                End If
            End If
        Next j
    Next i
    
    doc.Regen acAllViewports
End Sub

' Helper function: checks if a point is inside a circle
Function IsPointInsideCircle(circ As AcadCircle, pt As Variant) As Boolean
    Dim dx As Double, dy As Double, dz As Double
    dx = pt(0) - circ.Center(0)
    dy = pt(1) - circ.Center(1)
    dz = pt(2) - circ.Center(2)
    IsPointInsideCircle = Sqr(dx * dx + dy * dy + dz * dz) < circ.Radius
End Function
