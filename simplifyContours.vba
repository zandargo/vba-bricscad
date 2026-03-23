Option Explicit

' Converts a selected chain of lines, polylines and/or arcs into a spline,
' deleting the original entities. Runs in a loop until the user presses ESC
' or cancels the selection.

' Number of intermediate points used to sample each arc segment
Private Const ARC_SAMPLE_STEPS As Long = 16

' Snap tolerance for chain detection (drawing units)
Private Const CHAIN_TOL As Double = 0.01

Private Type Point2D
    X As Double
    Y As Double
End Type

' Represents a segment extracted from any entity type (line, arc, polyline).
' StartPt / EndPt are the physical endpoints; Points holds the full sampled
' point sequence from StartPt to EndPt.
Private Type Segment
    StartPt  As Point2D
    EndPt    As Point2D
    PtCount  As Long          ' number of entries in Pts()
    Pts(0 To 512) As Point2D  ' sampled interior + endpoints
    EntIdx   As Long          ' index in the original ss for deletion
End Type

Public Sub SimplifySelectedContours()
    Dim doc As Object
    Set doc = ThisDrawing

    Dim wasFormVisible As Boolean
    On Error Resume Next
    wasFormVisible = formPerfisul01.Visible
    If wasFormVisible Then formPerfisul01.Hide
    On Error GoTo SafeExit

    doc.Utility.Prompt "Chain-to-Spline: area-select a chain of lines/polylines/arcs, then press ENTER. Press ESC to exit." & vbCrLf

    Dim ss As Object
    Set ss = PrepareSelectionSet(doc, "SIMP_CHAIN")
    If ss Is Nothing Then Exit Sub

    Do
        ss.Clear

        doc.Utility.Prompt "Select chain (window/crossing) then ENTER, or ESC to exit: "

        On Error Resume Next
        ss.SelectOnScreen
        Dim selErr As Long: selErr = Err.Number
        On Error GoTo 0

        If selErr <> 0 Or ss.Count = 0 Then
            doc.Utility.Prompt "Exit." & vbCrLf
            Exit Do
        End If

        Dim converted As Long
        converted = ConvertChainToSpline(doc, ss)

        If converted = 0 Then
            doc.Utility.Prompt "No connected chain found in selection." & vbCrLf
        Else
            doc.Utility.Prompt converted & " spline(s) drawn." & vbCrLf
        End If
    Loop

SafeExit:
    On Error Resume Next
    If wasFormVisible Then formPerfisul01.Show
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Chain detection and spline conversion
' ---------------------------------------------------------------------------

' Convert selection to chains and draw one spline per chain.
Private Function ConvertChainToSpline(ByVal doc As Object, _
                                      ByVal ss As Object) As Long
    ' 1. Collect segments from all supported entities
    Dim segs() As Segment
    Dim segCount As Long
    ReDim segs(0 To ss.Count - 1)  ' upper bound; may not all be used

    Dim i As Long
    For i = 0 To ss.Count - 1
        Dim ent As Object
        Set ent = ss.Item(i)
        If ExtractSegment(ent, i, segs(segCount)) Then
            segCount = segCount + 1
        End If
    Next i

    If segCount = 0 Then Exit Function

    ' 2. Chain segments into ordered paths
    Dim used() As Boolean
    ReDim used(0 To segCount - 1)

    Dim splineCount As Long

    Dim s As Long
    For s = 0 To segCount - 1
        If used(s) Then GoTo NextSeg

        ' Start a new chain from this segment
        Dim chain() As Point2D
        Dim chainLen As Long
        Dim usedInChain() As Long
        ReDim usedInChain(0 To segCount - 1)
        Dim usedInChainCount As Long

        chain = CopySegPoints(segs(s))
        chainLen = segs(s).PtCount
        used(s) = True
        usedInChain(0) = s
        usedInChainCount = 1

        ' Grow chain by connecting matching endpoints
        Dim extended As Boolean
        Do
            extended = False
            Dim j As Long
            For j = 0 To segCount - 1
                If used(j) Then GoTo NextSegJ

                ' Try attaching segment j to either end of the chain
                If PointsAreEqual(chain(chainLen - 1), segs(j).StartPt, CHAIN_TOL) Then
                    ' Append seg j forward (skip duplicate junction point)
                    chain = AppendSegPoints(chain, chainLen, segs(j), True)
                    chainLen = chainLen + segs(j).PtCount - 1
                    used(j) = True
                    usedInChain(usedInChainCount) = j
                    usedInChainCount = usedInChainCount + 1
                    extended = True
                    GoTo RestartChain
                ElseIf PointsAreEqual(chain(chainLen - 1), segs(j).EndPt, CHAIN_TOL) Then
                    ' Append seg j reversed
                    chain = AppendSegPointsRev(chain, chainLen, segs(j), True)
                    chainLen = chainLen + segs(j).PtCount - 1
                    used(j) = True
                    usedInChain(usedInChainCount) = j
                    usedInChainCount = usedInChainCount + 1
                    extended = True
                    GoTo RestartChain
                ElseIf PointsAreEqual(chain(0), segs(j).EndPt, CHAIN_TOL) Then
                    ' Prepend seg j forward (its end connects to chain start)
                    chain = PrependSegPoints(chain, chainLen, segs(j), True)
                    chainLen = chainLen + segs(j).PtCount - 1
                    used(j) = True
                    usedInChain(usedInChainCount) = j
                    usedInChainCount = usedInChainCount + 1
                    extended = True
                    GoTo RestartChain
                ElseIf PointsAreEqual(chain(0), segs(j).StartPt, CHAIN_TOL) Then
                    ' Prepend seg j reversed
                    chain = PrependSegPointsRev(chain, chainLen, segs(j), True)
                    chainLen = chainLen + segs(j).PtCount - 1
                    used(j) = True
                    usedInChain(usedInChainCount) = j
                    usedInChainCount = usedInChainCount + 1
                    extended = True
                    GoTo RestartChain
                End If
NextSegJ:
            Next j
RestartChain:
        Loop While extended

        ' Need at least 2 distinct points for a spline
        If chainLen < 2 Then GoTo NextSeg

        ' Remove duplicate points from chain
        Dim cleanChain() As Point2D
        Dim cleanLen As Long
        CleanDuplicates chain, chainLen, cleanChain, cleanLen

        If cleanLen < 2 Then GoTo NextSeg

        ' 3. Draw spline through chain points
        Dim splineCoords() As Double
        ReDim splineCoords(0 To cleanLen * 3 - 1)
        Dim k As Long
        For k = 0 To cleanLen - 1
            splineCoords(k * 3)     = cleanChain(k).X
            splineCoords(k * 3 + 1) = cleanChain(k).Y
            splineCoords(k * 3 + 2) = 0#
        Next k

        Dim startTan(0 To 2) As Double
        Dim endTan(0 To 2) As Double

        ' Tangents at endpoints: direction of first/last segment
        startTan(0) = cleanChain(1).X - cleanChain(0).X
        startTan(1) = cleanChain(1).Y - cleanChain(0).Y
        startTan(2) = 0#
        endTan(0) = cleanChain(cleanLen - 1).X - cleanChain(cleanLen - 2).X
        endTan(1) = cleanChain(cleanLen - 1).Y - cleanChain(cleanLen - 2).Y
        endTan(2) = 0#

        On Error Resume Next
        Dim spl As Object
        Set spl = doc.ModelSpace.AddSpline(splineCoords, startTan, endTan)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextSeg
        End If
        On Error GoTo 0

        spl.Update
        splineCount = splineCount + 1

        ' 4. Delete original entities that were chained
        Dim d As Long
        For d = 0 To usedInChainCount - 1
            Dim entIdx As Long
            entIdx = segs(usedInChain(d)).EntIdx
            On Error Resume Next
            ss.Item(entIdx).Delete
            On Error GoTo 0
        Next d

NextSeg:
    Next s

    ConvertChainToSpline = splineCount
End Function

' ---- Segment extraction ----------------------------------------------------

Private Function ExtractSegment(ByVal ent As Object, ByVal ssIdx As Long, _
                                ByRef seg As Segment) As Boolean
    Dim t As String
    t = TypeName(ent)

    Select Case t
        Case "IAcadLine", "AcadLine"
            seg.StartPt = ToPoint2D(ent.StartPoint)
            seg.EndPt   = ToPoint2D(ent.EndPoint)
            seg.PtCount = 2
            seg.Pts(0)  = seg.StartPt
            seg.Pts(1)  = seg.EndPt
            seg.EntIdx  = ssIdx
            ExtractSegment = True

        Case "IAcadArc", "AcadArc"
            ExtractSegment = SampleArc(ent, ssIdx, seg)

        Case "IAcadLWPolyline", "AcadLWPolyline"
            ExtractSegment = SampleLWPolyline(ent, ssIdx, seg)

        Case "IAcadPolyline", "AcadPolyline"
            ExtractSegment = Sample3DPolyline(ent, ssIdx, seg)

        Case Else
            ExtractSegment = False
    End Select
End Function

Private Function SampleArc(ByVal arc As Object, ByVal ssIdx As Long, _
                            ByRef seg As Segment) As Boolean
    Dim startAng As Double: startAng = arc.StartAngle
    Dim endAng   As Double: endAng   = arc.EndAngle
    Dim cpt As Variant:     cpt = arc.Center
    Dim cx As Double:       cx = cpt(0)
    Dim cy As Double:       cy = cpt(1)
    Dim r  As Double:       r  = arc.Radius

    ' Normalise so we sweep CCW from startAng to endAng
    Dim span As Double
    span = endAng - startAng
    If span <= 0# Then span = span + 2# * 3.14159265358979#

    Dim steps As Long: steps = ARC_SAMPLE_STEPS
    Dim n As Long:     n = steps + 1          ' number of points
    If n > 512 Then n = 512

    Dim i As Long
    For i = 0 To n - 1
        Dim ang As Double
        ang = startAng + span * i / (n - 1)
        seg.Pts(i).X = cx + r * Cos(ang)
        seg.Pts(i).Y = cy + r * Sin(ang)
    Next i

    seg.PtCount  = n
    seg.StartPt  = seg.Pts(0)
    seg.EndPt    = seg.Pts(n - 1)
    seg.EntIdx   = ssIdx
    SampleArc = True
End Function

Private Function SampleLWPolyline(ByVal pl As Object, ByVal ssIdx As Long, _
                                  ByRef seg As Segment) As Boolean
    Dim arr As Variant: arr = pl.Coordinates
    Dim total As Long:  total = (UBound(arr) + 1) \ 2
    If total < 2 Then Exit Function
    If total > 513 Then total = 513

    Dim i As Long
    For i = 0 To total - 1
        seg.Pts(i).X = arr(2 * i)
        seg.Pts(i).Y = arr(2 * i + 1)
    Next i

    seg.PtCount = total
    seg.StartPt = seg.Pts(0)
    seg.EndPt   = seg.Pts(total - 1)
    seg.EntIdx  = ssIdx
    SampleLWPolyline = True
End Function

Private Function Sample3DPolyline(ByVal pl As Object, ByVal ssIdx As Long, _
                                  ByRef seg As Segment) As Boolean
    Dim arr As Variant: arr = pl.Coordinates
    Dim total As Long:  total = (UBound(arr) + 1) \ 3
    If total < 2 Then Exit Function
    If total > 513 Then total = 513

    Dim i As Long
    For i = 0 To total - 1
        seg.Pts(i).X = arr(3 * i)
        seg.Pts(i).Y = arr(3 * i + 1)
    Next i

    seg.PtCount = total
    seg.StartPt = seg.Pts(0)
    seg.EndPt   = seg.Pts(total - 1)
    seg.EntIdx  = ssIdx
    Sample3DPolyline = True
End Function

' ---- Chain point helpers ----------------------------------------------------

Private Function CopySegPoints(ByRef seg As Segment) As Point2D()
    Dim out() As Point2D
    ReDim out(0 To seg.PtCount - 1)
    Dim i As Long
    For i = 0 To seg.PtCount - 1
        out(i) = seg.Pts(i)
    Next i
    CopySegPoints = out
End Function

' Append seg points to existing chain (skip first point if skipFirst = True to
' avoid duplicating the junction).
Private Function AppendSegPoints(ByRef chain() As Point2D, ByVal chainLen As Long, _
                                 ByRef seg As Segment, ByVal skipFirst As Boolean) As Point2D()
    Dim startI As Long: startI = IIf(skipFirst, 1, 0)
    Dim addCount As Long: addCount = seg.PtCount - startI
    Dim newLen As Long: newLen = chainLen + addCount
    ReDim Preserve chain(0 To newLen - 1)
    Dim i As Long
    For i = 0 To addCount - 1
        chain(chainLen + i) = seg.Pts(startI + i)
    Next i
    AppendSegPoints = chain
End Function

Private Function AppendSegPointsRev(ByRef chain() As Point2D, ByVal chainLen As Long, _
                                    ByRef seg As Segment, ByVal skipFirst As Boolean) As Point2D()
    ' Append seg in reverse order (seg.EndPt first, seg.StartPt last)
    ' skipFirst: skip seg.EndPt (the junction already in chain)
    Dim startI As Long: startI = IIf(skipFirst, seg.PtCount - 2, seg.PtCount - 1)
    Dim addCount As Long: addCount = startI + 1
    Dim newLen As Long: newLen = chainLen + addCount
    ReDim Preserve chain(0 To newLen - 1)
    Dim i As Long
    For i = 0 To addCount - 1
        chain(chainLen + i) = seg.Pts(startI - i)
    Next i
    AppendSegPointsRev = chain
End Function

Private Function PrependSegPoints(ByRef chain() As Point2D, ByVal chainLen As Long, _
                                  ByRef seg As Segment, ByVal skipFirst As Boolean) As Point2D()
    ' Prepend seg (forward) before chain; seg.EndPt connects to chain(0)
    ' skipFirst: skip seg.EndPt
    Dim endI As Long: endI = IIf(skipFirst, seg.PtCount - 2, seg.PtCount - 1)
    Dim addCount As Long: addCount = endI + 1
    Dim newLen As Long: newLen = chainLen + addCount
    Dim out() As Point2D
    ReDim out(0 To newLen - 1)
    Dim i As Long
    For i = 0 To addCount - 1
        out(i) = seg.Pts(i)
    Next i
    For i = 0 To chainLen - 1
        out(addCount + i) = chain(i)
    Next i
    PrependSegPoints = out
End Function

Private Function PrependSegPointsRev(ByRef chain() As Point2D, ByVal chainLen As Long, _
                                     ByRef seg As Segment, ByVal skipFirst As Boolean) As Point2D()
    ' Prepend seg reversed before chain; seg.StartPt connects to chain(0)
    Dim startI As Long: startI = IIf(skipFirst, 1, 0)
    Dim addCount As Long: addCount = seg.PtCount - startI
    Dim newLen As Long: newLen = chainLen + addCount
    Dim out() As Point2D
    ReDim out(0 To newLen - 1)
    Dim i As Long
    For i = 0 To addCount - 1
        out(i) = seg.Pts(seg.PtCount - 1 - i)
    Next i
    For i = 0 To chainLen - 1
        out(addCount + i) = chain(i)
    Next i
    PrependSegPointsRev = out
End Function

Private Sub CleanDuplicates(ByRef pts() As Point2D, ByVal n As Long, _
                            ByRef out() As Point2D, ByRef outLen As Long)
    ReDim out(0 To n - 1)
    Dim i As Long
    out(0) = pts(0)
    outLen = 1
    For i = 1 To n - 1
        If Not PointsAreEqual(pts(i), out(outLen - 1), CHAIN_TOL) Then
            out(outLen) = pts(i)
            outLen = outLen + 1
        End If
    Next i
End Sub

' ---- Utility ----------------------------------------------------------------

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

Private Function PrepareSelectionSet(ByVal doc As Object, ByVal name As String) As Object
    On Error Resume Next
    Dim ss As Object
    Set ss = doc.SelectionSets.Item(name)
    If Not ss Is Nothing Then ss.Delete
    On Error GoTo 0

    Set ss = doc.SelectionSets.Add(name)
    Set PrepareSelectionSet = ss
End Function
