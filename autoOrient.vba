Option Explicit

' Auto-orient a selection by finding the best rotation angle (0-180 degrees).
' Call from form with: AutoOrientFromForm(Me)
' Call standalone with: AutoOrient()

Private Type Point2D
    x As Double
    y As Double
End Type

Public Sub AutoOrient()
    AutoOrientCore Nothing
End Sub

Public Sub AutoOrientFromForm(frm As Object)
    AutoOrientCore frm
End Sub

Private Sub AutoOrientCore(frm As Object)
    Dim doc As AcadDocument
    Set doc = ThisDrawing
    
    Dim wasVisible As Boolean: wasVisible = False
    If Not frm Is Nothing Then
        On Error Resume Next
        wasVisible = frm.Visible
        frm.Hide
        On Error GoTo 0
    End If
    
    On Error GoTo ErrorHandler
    
    ' Prompt user to select entities
    Dim ss As AcadSelectionSet
    Set ss = doc.SelectionSets.Add("AUTOORIENT_SS")
    
    doc.Utility.Prompt vbCr & "Selecione os objetos: " & vbCr
    ss.SelectOnScreen
    
    If ss.Count = 0 Then
        doc.Utility.Prompt "Nenhum objeto selecionado." & vbCr
        ss.Delete
        GoTo Cleanup
    End If
    
    doc.StartUndoMark
    
    ' Get bounding box of current selection
    Dim minPt As Variant, maxPt As Variant
    GetSelectionBounds ss, minPt, maxPt
    
    Dim centerPt(0 To 2) As Double
    centerPt(0) = (minPt(0) + maxPt(0)) / 2
    centerPt(1) = (minPt(1) + maxPt(1)) / 2
    centerPt(2) = 0
    
    ' Find best rotation angle
    Dim bestAngle As Double
    Dim bestHeight As Double
    bestAngle = FindBestRotationAngleInternal(ss, centerPt, bestHeight)
    
    Dim degAngle As Double
    degAngle = bestAngle * 180 / 3.14159265358979
    
    doc.Utility.Prompt "Melhor angulo: " & Format(degAngle, "0.00") & " graus" & vbCr
    
    ' Apply rotation if angle is significant
    If Abs(bestAngle) > 0.001 Then
        RotateSelectedEntities ss, centerPt, bestAngle
        MsgBox "Rotacionado em " & Format(degAngle, "0.00") & " graus.", vbInformation
    Else
        MsgBox "Nenhuma rotacao necessaria (angulo ~0).", vbInformation
    End If
    
    doc.EndUndoMark
    ss.Delete
    
Cleanup:
    If Not frm Is Nothing And wasVisible Then
        On Error Resume Next
        frm.Show
        On Error GoTo 0
    End If
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not ss Is Nothing Then ss.Delete
    MsgBox "Erro: " & Err.Description, vbCritical
    If Not frm Is Nothing And wasVisible Then frm.Show
End Sub

' Get bounding box of all entities in selection set
Private Sub GetSelectionBounds(ss As AcadSelectionSet, ByRef minPt As Variant, ByRef maxPt As Variant)
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double
    Dim first As Boolean: first = True
    
    Dim ent As AcadEntity
    Dim entMinPt As Variant, entMaxPt As Variant
    
    For Each ent In ss
        On Error Resume Next
        ent.GetBoundingBox entMinPt, entMaxPt
        
        If Not first Then
            If entMinPt(0) < minX Then minX = entMinPt(0)
            If entMinPt(1) < minY Then minY = entMinPt(1)
            If entMaxPt(0) > maxX Then maxX = entMaxPt(0)
            If entMaxPt(1) > maxY Then maxY = entMaxPt(1)
        Else
            minX = entMinPt(0)
            minY = entMinPt(1)
            maxX = entMaxPt(0)
            maxY = entMaxPt(1)
            first = False
        End If
        On Error GoTo 0
    Next ent
    
    minPt = Array(minX, minY, 0)
    maxPt = Array(maxX, maxY, 0)
End Sub

' Find best rotation angle by testing increments from 0 to 180 degrees
Private Function FindBestRotationAngleInternal(ss As AcadSelectionSet, centerPt() As Double, ByRef heightOut As Double) As Double
    Const STEP_DEG As Double = 1
    Const PI As Double = 3.14159265358979
    
    ' Gather all sampling points relative to centerPt ONCE
    Dim points() As Point2D
    Dim numPoints As Long
    numPoints = CollectSamplingPoints(ss, centerPt, points)
    
    Dim bestAngle As Double: bestAngle = 0
    Dim bestHeight As Double: bestHeight = 1E+30
    Dim bestAspect As Double: bestAspect = 0
    
    Dim angle As Double
    Dim deg As Double
    Dim width As Double, height As Double, aspect As Double
    
    ' Test angles from 0 to 180 degrees
    For deg = 0 To 180 Step STEP_DEG
        angle = deg * PI / 180
        
        ' Get bounds from pre-collected points rotated by angle
        GetRotatedBoundsFromPoints points, numPoints, angle, width, height
        
        If height > 0 Then
            aspect = width / height
        Else
            aspect = 0
        End If
        
        ' Prefer smaller height; use aspect ratio as tiebreaker
        If height < bestHeight - 0.001 Or (Abs(height - bestHeight) < 0.001 And aspect > bestAspect) Then
            bestHeight = height
            bestAspect = aspect
            bestAngle = angle
        End If
    Next deg
    
    heightOut = bestHeight
    FindBestRotationAngleInternal = bestAngle
End Function

Private Function CollectSamplingPoints(ss As AcadSelectionSet, centerPt() As Double, ByRef pointsOut() As Point2D) As Long
    Dim ent As AcadEntity
    Dim count As Long: count = 0
    ReDim pointsOut(0 To 1000) As Point2D
    
    For Each ent In ss
        Dim objName As String
        objName = UCase$(ent.ObjectName)
        
        ' For Polylines, explode to get accurate geometry (lines and arcs)
        If InStr(1, objName, "POLYLINE", vbTextCompare) > 0 Then
            CollectPolylinePoints ent, centerPt, pointsOut, count
        ElseIf objName = "ACDBLINE" Or objName = "ACADLINE" Then
            CollectLinePoints ent, centerPt, pointsOut, count
        ElseIf objName = "ACDBARC" Or objName = "ACADARC" Then
            CollectArcPoints ent, centerPt, pointsOut, count
        Else
            ' Fallback to bounding box corners for other entities
            CollectBoundingBoxPoints ent, centerPt, pointsOut, count
        End If
    Next ent
    
    CollectSamplingPoints = count
End Function

Private Sub AddPoint(x As Double, y As Double, ByRef points() As Point2D, ByRef count As Long)
    If count > UBound(points) Then
        ReDim Preserve points(0 To UBound(points) * 2) As Point2D
    End If
    points(count).x = x
    points(count).y = y
    count = count + 1
End Sub

Private Sub CollectPolylinePoints(ent As AcadEntity, centerPt() As Double, ByRef points() As Point2D, ByRef count As Long)
    On Error Resume Next
    Dim exploded As Variant
    exploded = ent.Explode
    
    If Err.Number <> 0 Or IsEmpty(exploded) Then
        Err.Clear
        CollectBoundingBoxPoints ent, centerPt, points, count
        Exit Sub
    End If
    
    Dim i As Long
    Dim subEnt As AcadEntity
    For i = LBound(exploded) To UBound(exploded)
        Set subEnt = exploded(i)
        
        Dim subName As String
        subName = UCase$(subEnt.ObjectName)
        
        If InStr(1, subName, "LINE", vbTextCompare) > 0 Then
            CollectLinePoints subEnt, centerPt, points, count
        ElseIf InStr(1, subName, "ARC", vbTextCompare) > 0 Then
            CollectArcPoints subEnt, centerPt, points, count
        Else
            CollectBoundingBoxPoints subEnt, centerPt, points, count
        End If
        
        subEnt.Delete 
    Next i
    On Error GoTo 0
End Sub

Private Sub CollectLinePoints(lineEnt As AcadEntity, centerPt() As Double, ByRef points() As Point2D, ByRef count As Long)
    On Error Resume Next
    Dim startPt As Variant, endPt As Variant
    startPt = lineEnt.StartPoint
    endPt = lineEnt.EndPoint
    
    AddPoint startPt(0) - centerPt(0), startPt(1) - centerPt(1), points, count
    AddPoint endPt(0) - centerPt(0), endPt(1) - centerPt(1), points, count
    On Error GoTo 0
End Sub

Private Sub CollectArcPoints(arcEnt As AcadEntity, centerPt() As Double, ByRef points() As Point2D, ByRef count As Long)
    On Error Resume Next
    Dim startPt As Variant, endPt As Variant
    startPt = arcEnt.StartPoint
    endPt = arcEnt.EndPoint
    
    AddPoint startPt(0) - centerPt(0), startPt(1) - centerPt(1), points, count
    AddPoint endPt(0) - centerPt(0), endPt(1) - centerPt(1), points, count
    
    ' Sample midpoint of arc
    Dim radius As Double, center As Variant, startAngle As Double, endAngle As Double
    radius = arcEnt.radius
    center = arcEnt.center
    startAngle = arcEnt.startAngle
    endAngle = arcEnt.endAngle
    
    ' Normalize angles for arc span calculation
    ' BricsCAD/AutoCAD arcs are CCW
    
    Dim angleDiff As Double
    angleDiff = endAngle - startAngle
    If angleDiff <= 0 Then angleDiff = angleDiff + 6.28318530717959
    
    Dim midAngle As Double
    midAngle = startAngle + (angleDiff / 2)
    
    Dim midX As Double, midY As Double
    midX = center(0) + radius * Cos(midAngle)
    midY = center(1) + radius * Sin(midAngle)
    
    AddPoint midX - centerPt(0), midY - centerPt(1), points, count
    On Error GoTo 0
End Sub

Private Sub CollectBoundingBoxPoints(ent As AcadEntity, centerPt() As Double, ByRef points() As Point2D, ByRef count As Long)
    On Error Resume Next
    Dim minPt As Variant, maxPt As Variant
    ent.GetBoundingBox minPt, maxPt
    
    If Err.Number = 0 Then
        AddPoint minPt(0) - centerPt(0), minPt(1) - centerPt(1), points, count
        AddPoint maxPt(0) - centerPt(0), minPt(1) - centerPt(1), points, count
        AddPoint maxPt(0) - centerPt(0), maxPt(1) - centerPt(1), points, count
        AddPoint minPt(0) - centerPt(0), maxPt(1) - centerPt(1), points, count
    End If
    On Error GoTo 0
End Sub

Private Sub GetRotatedBoundsFromPoints(points() As Point2D, count As Long, angle As Double, _
    ByRef widthOut As Double, ByRef heightOut As Double)
    
    Const LARGE As Double = 1E+30
    Dim minX As Double: minX = LARGE
    Dim minY As Double: minY = LARGE
    Dim maxX As Double: maxX = -LARGE
    Dim maxY As Double: maxY = -LARGE
    
    Dim cosA As Double: cosA = Cos(angle)
    Dim sinA As Double: sinA = Sin(angle)
    Dim i As Long
    Dim rx As Double, ry As Double
    
    For i = 0 To count - 1
        rx = points(i).x * cosA - points(i).y * sinA
        ry = points(i).x * sinA + points(i).y * cosA
        
        If rx < minX Then minX = rx
        If rx > maxX Then maxX = rx
        If ry < minY Then minY = ry
        If ry > maxY Then maxY = ry
    Next i
    
    If minX > maxX Then ' No points processed
        widthOut = 0
        heightOut = 0
    Else
        widthOut = maxX - minX
        heightOut = maxY - minY
    End If
End Sub

' Rotate all entities in the selection set
Private Sub RotateSelectedEntities(ss As AcadSelectionSet, centerPt() As Double, angle As Double)
    Dim ent As AcadEntity
    Dim errCount As Long: errCount = 0
    
    For Each ent In ss
        On Error Resume Next
        ent.Rotate centerPt, angle
        If Err.Number <> 0 Then
            errCount = errCount + 1
            Err.Clear
        End If
        On Error GoTo 0
    Next ent
    
    If errCount > 0 Then
        MsgBox "Aviso: " & errCount & " entidade(s) nao pudo ser rotacionada(s).", vbExclamation
    End If
End Sub
