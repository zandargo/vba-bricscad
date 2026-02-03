Option Explicit

' Simplifies selected line/polyline contours using a Douglas-Peucker pass
' plus a near-collinear cleanup. Produces simplified LW polylines on a
' dedicated layer without deleting the originals.

' User-adjustable defaults
Private Const DEFAULT_TOLERANCE As Double = 0.5       ' Max deviation in drawing units
Private Const DEFAULT_ANGLE_TOL_DEG As Double = 5#    ' Max angle considered collinear

Private Type Point2D
    X As Double
    Y As Double
End Type

Private Type LineSegment
    A As Point2D
    B As Point2D
End Type

Public Sub SimplifySelectedContours()
    Dim doc As AcadDocument
    Set doc = ThisDrawing

    Dim tol As Double: tol = DEFAULT_TOLERANCE
    Dim angTol As Double: angTol = DEFAULT_ANGLE_TOL_DEG

    Dim ss As AcadSelectionSet
    Set ss = PrepareSelectionSet(doc, "SIMP_SIMPLIFY")
    If ss Is Nothing Then Exit Sub

    On Error Resume Next
    ss.SelectOnScreen
    If Err.Number <> 0 Then
        MsgBox "Selection cancelled.", vbInformation, "Simplify"
        Exit Sub
    End If
    On Error GoTo 0

    If ss.Count = 0 Then
        MsgBox "Nothing selected.", vbInformation, "Simplify"
        Exit Sub
    End If

    Dim simplifiedCount As Long
    simplifiedCount = ProcessSelection(doc, ss, tol, angTol)

    MsgBox "Simplified polylines created: " & simplifiedCount, vbInformation, "Simplify"
End Sub

Private Function ProcessSelection(ByVal doc As AcadDocument, ByVal ss As AcadSelectionSet, _
                                  ByVal tol As Double, ByVal angTolDeg As Double) As Long
    Dim lay As AcadLayer
    Set lay = EnsureLayer(doc, "SIMPLIFIED")

    Dim pathCoords As Collection ' each item is a Variant array of doubles
    Dim pathClosed As Collection ' each item is a Boolean
    Set pathCoords = New Collection
    Set pathClosed = New Collection

    Dim lineSegs() As LineSegment
    Dim lineCount As Long

    Dim idx As Long
    Dim ent As AcadEntity

    For idx = 0 To ss.Count - 1
        Set ent = ss.Item(idx)

        Dim pts() As Point2D
        Dim isClosed As Boolean

        If CollectVertices(ent, pts, isClosed) Then
            AddPathToCollections pts, isClosed, pathCoords, pathClosed
        ElseIf IsLineEntity(ent) Then
            Dim seg As LineSegment
            seg.A = ToPoint2D(ent.StartPoint)
            seg.B = ToPoint2D(ent.EndPoint)
            ReDim Preserve lineSegs(0 To lineCount)
            lineSegs(lineCount) = seg
            lineCount = lineCount + 1
        End If
    Next idx

    BuildPathsFromLines lineSegs, lineCount, tol, pathCoords, pathClosed

    Dim simplified As Long
    Dim listIdx As Long
    For listIdx = 1 To pathCoords.Count
        Dim ptsVar As Variant
        ptsVar = pathCoords.Item(listIdx)

        Dim ptsArr() As Point2D
        If Not VariantToPoints(ptsVar, ptsArr) Then GoTo ContinueLoop

        Dim closedFlag As Boolean
        closedFlag = pathClosed.Item(listIdx)

        Dim simplifiedPts As Variant
        simplifiedPts = SimplifyPath(ptsArr, closedFlag, tol, angTolDeg)
        If IsEmpty(simplifiedPts) Then GoTo ContinueLoop
        If UBound(simplifiedPts) < 1 Then GoTo ContinueLoop

        Dim pline As AcadLWPolyline
        Set pline = doc.ModelSpace.AddLightWeightPolyline(simplifiedPts)
        pline.Closed = closedFlag
        pline.Layer = lay.Name
        pline.Update

        simplified = simplified + 1
ContinueLoop:
    Next listIdx

    ProcessSelection = simplified
End Function

Private Function IsLineEntity(ByVal ent As AcadEntity) As Boolean
    Dim t As String
    t = TypeName(ent)
    IsLineEntity = (t = "IAcadLine" Or t = "AcadLine")
End Function

Private Function CollectVertices(ByVal ent As AcadEntity, ByRef pts() As Point2D, _
                                 ByRef isClosed As Boolean) As Boolean
    Dim t As String
    t = TypeName(ent)

    Select Case t
        Case "IAcadLine", "AcadLine"
            ReDim pts(0 To 1)
            pts(0) = ToPoint2D(ent.StartPoint)
            pts(1) = ToPoint2D(ent.EndPoint)
            isClosed = False
            CollectVertices = True
        Case "IAcadLWPolyline", "AcadLWPolyline"
            CollectVertices = ExtractPolylineVertices(ent, pts, isClosed)
        Case "IAcadPolyline", "AcadPolyline"
            CollectVertices = Extract3DPolylineVertices(ent, pts, isClosed)
        Case Else
            CollectVertices = False
    End Select
End Function

Private Function ExtractPolylineVertices(ByVal pl As AcadLWPolyline, _
                                         ByRef pts() As Point2D, _
                                         ByRef isClosed As Boolean) As Boolean
    Dim arr As Variant
    arr = pl.Coordinates
    Dim n As Long: n = UBound(arr)
    If (n + 1) Mod 2 <> 0 Then Exit Function

    Dim count As Long: count = (n + 1) \ 2
    ReDim pts(0 To count - 1)

    Dim i As Long
    For i = 0 To count - 1
        pts(i).X = arr(2 * i)
        pts(i).Y = arr(2 * i + 1)
    Next i

    isClosed = pl.Closed
    ExtractPolylineVertices = True
End Function

Private Function Extract3DPolylineVertices(ByVal pl As AcadPolyline, _
                                           ByRef pts() As Point2D, _
                                           ByRef isClosed As Boolean) As Boolean
    Dim i As Long
    Dim count As Long: count = pl.Coordinates.Count \ 3
    If count = 0 Then Exit Function

    ReDim pts(0 To count - 1)
    Dim arr As Variant: arr = pl.Coordinates

    For i = 0 To count - 1
        pts(i).X = arr(3 * i)
        pts(i).Y = arr(3 * i + 1)
    Next i

    isClosed = pl.Closed
    Extract3DPolylineVertices = True
End Function

Private Function SimplifyPath(ByRef pts() As Point2D, ByVal isClosed As Boolean, _
                              ByVal tol As Double, ByVal angTolDeg As Double) As Variant
    Dim working() As Point2D
    working = pts

    ' If closed, drop duplicate terminal point if present to simplify logic
    If isClosed And PointsAreEqual(working(LBound(working)), working(UBound(working)), tol / 10#) Then
        ReDim Preserve working(LBound(working) To UBound(working) - 1)
    End If

    If UBound(working) - LBound(working) + 1 < 2 Then Exit Function

    Dim keep() As Boolean
    ReDim keep(LBound(working) To UBound(working))

    keep(LBound(working)) = True
    keep(UBound(working)) = True
    RDP working, LBound(working), UBound(working), tol, keep

    Dim reduced() As Point2D
    reduced = PointsByKeepMask(working, keep)

    reduced = RemoveShortSegments(reduced, isClosed, tol)
    reduced = RemoveNearCollinear(reduced, isClosed, angTolDeg)

    If isClosed Then
        reduced = EnsureClosure(reduced, tol)
    End If

    SimplifyPath = PointsToVariant(reduced)
End Function

Private Sub RDP(ByRef pts() As Point2D, ByVal firstIdx As Long, ByVal lastIdx As Long, _
                ByVal tol As Double, ByRef keep() As Boolean)
    Dim maxDist As Double
    Dim idxMax As Long
    Dim i As Long

    For i = firstIdx + 1 To lastIdx - 1
        Dim d As Double
        d = PerpDistance(pts(i), pts(firstIdx), pts(lastIdx))
        If d > maxDist Then
            maxDist = d
            idxMax = i
        End If
    Next i

    If maxDist > tol Then
        keep(idxMax) = True
        If idxMax - firstIdx > 1 Then RDP pts, firstIdx, idxMax, tol, keep
        If lastIdx - idxMax > 1 Then RDP pts, idxMax, lastIdx, tol, keep
    End If
End Sub

Private Function PerpDistance(ByRef p As Point2D, ByRef a As Point2D, ByRef b As Point2D) As Double
    Dim vx As Double, vy As Double
    vx = b.X - a.X
    vy = b.Y - a.Y

    Dim wx As Double, wy As Double
    wx = p.X - a.X
    wy = p.Y - a.Y

    Dim segLen2 As Double
    segLen2 = vx * vx + vy * vy
    If segLen2 = 0# Then
        PerpDistance = Sqr(wx * wx + wy * wy)
        Exit Function
    End If

    Dim proj As Double
    proj = (wx * vx + wy * vy) / segLen2
    If proj < 0# Then
        proj = 0#
    ElseIf proj > 1# Then
        proj = 1#
    End If

    Dim cx As Double, cy As Double
    cx = a.X + proj * vx
    cy = a.Y + proj * vy

    Dim dx As Double, dy As Double
    dx = p.X - cx
    dy = p.Y - cy

    PerpDistance = Sqr(dx * dx + dy * dy)
End Function

Private Function PointsByKeepMask(ByRef pts() As Point2D, ByRef keep() As Boolean) As Point2D()
    Dim i As Long, c As Long
    For i = LBound(keep) To UBound(keep)
        If keep(i) Then c = c + 1
    Next i

    Dim out() As Point2D
    ReDim out(0 To c - 1)

    Dim idx As Long
    For i = LBound(keep) To UBound(keep)
        If keep(i) Then
            out(idx) = pts(i)
            idx = idx + 1
        End If
    Next i

    PointsByKeepMask = out
End Function

Private Function RemoveNearCollinear(ByRef pts() As Point2D, ByVal isClosed As Boolean, _
                                     ByVal angTolDeg As Double) As Point2D()
    Dim n As Long
    n = UBound(pts) - LBound(pts) + 1
    If n < 3 Then
        RemoveNearCollinear = pts
        Exit Function
    End If

    Dim changed As Boolean
    Dim threshold As Double
    threshold = Cos(angTolDeg * 3.14159265358979# / 180#)

    Dim work() As Point2D
    work = pts

    Do
        changed = False
        n = UBound(work) - LBound(work) + 1
        If n < 3 Then Exit Do

        Dim i As Long
        For i = 0 To n - 1
            Dim prevIdx As Long, nextIdx As Long
            prevIdx = (i - 1 + n) Mod n
            nextIdx = (i + 1) Mod n

            If (Not isClosed) And (i = 0 Or i = n - 1) Then GoTo ContinueLoop

            Dim v1x As Double, v1y As Double, v2x As Double, v2y As Double
            v1x = work(i).X - work(prevIdx).X
            v1y = work(i).Y - work(prevIdx).Y
            v2x = work(nextIdx).X - work(i).X
            v2y = work(nextIdx).Y - work(i).Y

            Dim len1 As Double, len2 As Double
            len1 = Sqr(v1x * v1x + v1y * v1y)
            len2 = Sqr(v2x * v2x + v2y * v2y)
            If len1 = 0# Or len2 = 0# Then GoTo ContinueLoop

            Dim dot As Double
            dot = (v1x * v2x + v1y * v2y) / (len1 * len2)

            If dot >= threshold Then
                work = RemoveIndex(work, i)
                changed = True
                Exit For
            End If
ContinueLoop:
        Next i
    Loop While changed

    RemoveNearCollinear = work
End Function

Private Function RemoveShortSegments(ByRef pts() As Point2D, ByVal isClosed As Boolean, _
                                     ByVal tol As Double) As Point2D()
    Dim n As Long
    n = UBound(pts) - LBound(pts) + 1
    If n < 3 Then
        RemoveShortSegments = pts
        Exit Function
    End If

    Dim work() As Point2D
    work = pts

    Dim changed As Boolean
    Dim tol2 As Double
    tol2 = tol * tol

    Do
        changed = False
        n = UBound(work) - LBound(work) + 1
        If n < 3 Then Exit Do

        Dim i As Long
        For i = 0 To n - 1
            Dim prevIdx As Long, nextIdx As Long
            prevIdx = (i - 1 + n) Mod n
            nextIdx = (i + 1) Mod n

            If (Not isClosed) And (i = 0 Or i = n - 1) Then GoTo ContinueLoop

            Dim dx As Double, dy As Double
            dx = work(nextIdx).X - work(i).X
            dy = work(nextIdx).Y - work(i).Y
            If dx * dx + dy * dy < tol2 Then
                work = RemoveIndex(work, nextIdx)
                changed = True
                Exit For
            End If
ContinueLoop:
        Next i
    Loop While changed

    RemoveShortSegments = work
End Function

Private Function RemoveIndex(ByRef pts() As Point2D, ByVal idx As Long) As Point2D()
    Dim n As Long
    n = UBound(pts) - LBound(pts) + 1
    If n <= 1 Then
        RemoveIndex = pts
        Exit Function
    End If

    Dim out() As Point2D
    ReDim out(0 To n - 2)

    Dim i As Long, j As Long
    For i = 0 To n - 1
        If i <> idx Then
            out(j) = pts(i)
            j = j + 1
        End If
    Next i

    RemoveIndex = out
End Function

Private Function EnsureClosure(ByRef pts() As Point2D, ByVal tol As Double) As Point2D()
    Dim n As Long
    n = UBound(pts) - LBound(pts) + 1
    If n < 2 Then
        EnsureClosure = pts
        Exit Function
    End If

    If Not PointsAreEqual(pts(LBound(pts)), pts(UBound(pts)), tol) Then
        ReDim Preserve pts(LBound(pts) To UBound(pts) + 1)
        pts(UBound(pts)) = pts(LBound(pts))
    End If

    EnsureClosure = pts
End Function

Private Sub BuildPathsFromLines(ByRef segs() As LineSegment, ByVal count As Long, _
                                ByVal tol As Double, ByRef pathCoords As Collection, _
                                ByRef pathClosed As Collection)
    If count = 0 Then Exit Sub

    Dim used() As Boolean
    ReDim used(0 To count - 1)

    Dim i As Long
    For i = 0 To count - 1
        If used(i) Then GoTo NextI

        Dim pts() As Point2D
        ReDim pts(0 To 1)
        pts(0) = segs(i).A
        pts(1) = segs(i).B
        used(i) = True

        Dim extended As Boolean
        Do
            extended = False
            Dim j As Long
            For j = 0 To count - 1
                If used(j) Then GoTo NextJ

                If PointsAreEqual(pts(LBound(pts)), segs(j).B, tol) Then
                    pts = PrependPoint(pts, segs(j).A)
                    used(j) = True
                    extended = True
                    GoTo RestartLoop
                ElseIf PointsAreEqual(pts(LBound(pts)), segs(j).A, tol) Then
                    pts = PrependPoint(pts, segs(j).B)
                    used(j) = True
                    extended = True
                    GoTo RestartLoop
                ElseIf PointsAreEqual(pts(UBound(pts)), segs(j).A, tol) Then
                    pts = AppendPoint(pts, segs(j).B)
                    used(j) = True
                    extended = True
                    GoTo RestartLoop
                ElseIf PointsAreEqual(pts(UBound(pts)), segs(j).B, tol) Then
                    pts = AppendPoint(pts, segs(j).A)
                    used(j) = True
                    extended = True
                    GoTo RestartLoop
                End If
NextJ:
            Next j
RestartLoop:
        Loop While extended

        Dim isClosed As Boolean
        isClosed = PointsAreEqual(pts(LBound(pts)), pts(UBound(pts)), tol)
        AddPathToCollections pts, isClosed, pathCoords, pathClosed
NextI:
    Next i

    End Sub

Private Function PrependPoint(ByRef pts() As Point2D, ByRef p As Point2D) As Point2D()
    Dim n As Long
    n = UBound(pts) - LBound(pts) + 1
    Dim out() As Point2D
    ReDim out(0 To n)

    out(0) = p
    Dim i As Long
    For i = 0 To n - 1
        out(i + 1) = pts(LBound(pts) + i)
    Next i

    PrependPoint = out
End Function

Private Function AppendPoint(ByRef pts() As Point2D, ByRef p As Point2D) As Point2D()
    Dim n As Long
    n = UBound(pts) - LBound(pts) + 1
    ReDim Preserve pts(LBound(pts) To UBound(pts) + 1)
    pts(UBound(pts)) = p
    AppendPoint = pts
End Function

Private Sub AddPathToCollections(ByRef pts() As Point2D, ByVal isClosed As Boolean, _
                                 ByRef pathCoords As Collection, ByRef pathClosed As Collection)
    Dim v As Variant
    v = PointsToVariant(pts)
    If Not IsEmpty(v) Then
        pathCoords.Add v
        pathClosed.Add isClosed
    End If
End Sub

Private Function VariantToPoints(ByVal v As Variant, ByRef pts() As Point2D) As Boolean
    If IsEmpty(v) Then Exit Function
    If Not IsArray(v) Then Exit Function
    Dim n As Long
    n = UBound(v)
    If (n + 1) Mod 2 <> 0 Then Exit Function

    Dim count As Long
    count = (n + 1) \ 2
    ReDim pts(0 To count - 1)

    Dim i As Long
    For i = 0 To count - 1
        pts(i).X = v(2 * i)
        pts(i).Y = v(2 * i + 1)
    Next i

    VariantToPoints = True
End Function

Private Function PointsToVariant(ByRef pts() As Point2D) As Variant
    Dim n As Long
    n = UBound(pts) - LBound(pts) + 1
    If n = 0 Then Exit Function

    Dim arr() As Double
    ReDim arr(0 To 2 * n - 1)

    Dim i As Long
    For i = 0 To n - 1
        arr(2 * i) = pts(i).X
        arr(2 * i + 1) = pts(i).Y
    Next i

    PointsToVariant = arr
End Function

Private Function ToPoint2D(ByVal arr As Variant) As Point2D
    ToPoint2D.X = arr(0)
    ToPoint2D.Y = arr(1)
End Function

Private Function PointsAreEqual(ByRef a As Point2D, ByRef b As Point2D, ByVal tol As Double) As Boolean
    Dim dx As Double, dy As Double
    dx = a.X - b.X
    dy = a.Y - b.Y
    PointsAreEqual = (dx * dx + dy * dy) <= tol * tol
End Function

Private Function PrepareSelectionSet(ByVal doc As AcadDocument, ByVal name As String) As AcadSelectionSet
    On Error Resume Next
    Dim ss As AcadSelectionSet
    Set ss = doc.SelectionSets.Item(name)
    If Not ss Is Nothing Then ss.Delete
    On Error GoTo 0

    Set ss = doc.SelectionSets.Add(name)
    Set PrepareSelectionSet = ss
End Function

Private Function EnsureLayer(ByVal doc As AcadDocument, ByVal layerName As String) As AcadLayer
    Dim lay As AcadLayer
    On Error Resume Next
    Set lay = doc.Layers.Item(layerName)
    On Error GoTo 0

    If lay Is Nothing Then
        Set lay = doc.Layers.Add(layerName)
    End If

    Set EnsureLayer = lay
End Function
