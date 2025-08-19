Public Sub FixSmallGaps()
    ' Procedure to find and fix small gaps between drawing elements
    ' Author: Created for BricsCAD VBA
    ' Date: August 2025
    
    Dim doc As AcadDocument
    Dim modelSpace As AcadModelSpace
    Dim entity As AcadEntity
    Dim i As Integer, j As Integer
    Dim minGap As Double, maxGap As Double
    Dim minGapMm As Double, maxGapMm As Double
    Dim response As String
    Dim userResponse As VbMsgBoxResult
    Dim unitFactor As Double
    Dim connectingLinesCount As Integer
    
    ' Initialize counters
    connectingLinesCount = 0
    
    ' Initialize document and model space
    Set doc = ThisDrawing
    Set modelSpace = doc.ModelSpace
    
    ' Get gap values from user (in mm)
    response = InputBox("Enter minimum gap value in mm:", "Minimum Gap", "0.00001")
    If response = "" Then Exit Sub
    ' Handle decimal separator issues (replace comma with period if needed)
    response = Replace(response, ",", ".")
    minGapMm = Val(response)
    ' MsgBox "Debug: minGapMm = " & minGapMm, vbInformation ' Debug line
    
    response = InputBox("Enter maximum gap value in mm:", "Maximum Gap", "0.05")
    If response = "" Then Exit Sub
    ' Handle decimal separator issues (replace comma with period if needed)
    response = Replace(response, ",", ".")
    maxGapMm = Val(response)
    ' MsgBox "Debug: maxGapMm = " & maxGapMm, vbInformation ' Debug line
    
    ' Convert from mm to drawing units
    ' Ask user about drawing units to ensure correct conversion
    Dim unitsResponse As String
    unitsResponse = InputBox("What are your drawing units?" & vbCrLf & _
                           "Enter 'mm' for millimeters" & vbCrLf & _
                           "Enter 'm' for meters" & vbCrLf & _
                           "Enter 'in' for inches", _
                           "Drawing Units", "mm")
    If unitsResponse = "" Then Exit Sub
    
    Select Case LCase(unitsResponse)
        Case "mm"
            unitFactor = 1 ' No conversion needed
        Case "m"
            unitFactor = 0.001 ' Convert mm to meters
        Case "in"
            unitFactor = 1 / 25.4 ' Convert mm to inches
        Case Else
            MsgBox "Invalid unit. Assuming meters.", vbExclamation
            unitFactor = 0.001
    End Select
    
    minGap = minGapMm * unitFactor
    maxGap = maxGapMm * unitFactor
    
    ' Debug info
    MsgBox "Converting " & minGapMm & "-" & maxGapMm & " mm to " & Format(minGap, "0.000000") & "-" & Format(maxGap, "0.000000") & " drawing units", vbInformation
    
    ' Validate input
    If minGap >= maxGap Then
        MsgBox "Minimum gap must be less than maximum gap!", vbExclamation
        Exit Sub
    End If
    
    ' Create collections to store line/arc/curve entities and their endpoints
    Dim entities As New Collection
    Dim endpoints As New Collection
    Dim entityTypes As New Collection
    
    ' Collect all lines, arcs, and curves with their endpoints
    ' MsgBox "Analyzing drawing elements...", vbInformation
    
    For i = 0 To modelSpace.Count - 1
        Set entity = modelSpace.Item(i)
        
        Select Case entity.ObjectName
            Case "AcDbLine"
                Dim line As AcadLine
                Set line = entity
                entities.Add line
                endpoints.Add Array(line.StartPoint, line.EndPoint)
                entityTypes.Add "Line"
                
            Case "AcDbArc"
                Dim arc As AcadArc
                Set arc = entity
                entities.Add arc
                ' Calculate arc endpoints from center, radius, and angles with high precision
                Dim arcStartPt(2) As Double, arcEndPt(2) As Double
                arcStartPt(0) = arc.center(0) + arc.radius * Cos(arc.StartAngle)
                arcStartPt(1) = arc.center(1) + arc.radius * Sin(arc.StartAngle)
                arcStartPt(2) = arc.center(2)
                arcEndPt(0) = arc.center(0) + arc.radius * Cos(arc.EndAngle)
                arcEndPt(1) = arc.center(1) + arc.radius * Sin(arc.EndAngle)
                arcEndPt(2) = arc.center(2)
                endpoints.Add Array(arcStartPt, arcEndPt)
                entityTypes.Add "Arc"
                
            Case "AcDbPolyline"
                Dim pline As AcadLWPolyline
                Set pline = entity
                entities.Add pline
                ' Get all coordinates and extract first and last points
                Dim coords As Variant
                coords = pline.Coordinates
                Dim startPt(2) As Double, endPt(2) As Double
                ' First vertex (coordinates are stored as X1,Y1,X2,Y2,...)
                startPt(0) = coords(0): startPt(1) = coords(1): startPt(2) = 0
                ' Last vertex
                Dim lastIndex As Integer
                lastIndex = UBound(coords) - 1
                endPt(0) = coords(lastIndex - 1): endPt(1) = coords(lastIndex): endPt(2) = 0
                endpoints.Add Array(startPt, endPt)
                entityTypes.Add "Polyline"
                
            Case "AcDbSpline"
                Dim spline As AcadSpline
                Set spline = entity
                entities.Add spline
                ' Safely attempt to read start/end points; some spline objects may not expose these properties
                On Error Resume Next
                Dim sStart As Variant, sEnd As Variant
                sStart = spline.StartPoint
                sEnd = spline.EndPoint
                If Err.Number <> 0 Then
                    Err.Clear
                    Dim bbMin(2) As Double, bbMax(2) As Double
                    ' Fall back to bounding box corners if StartPoint/EndPoint are not available
                    spline.GetBoundingBox bbMin, bbMax
                    sStart = bbMin
                    sEnd = bbMax
                End If
                On Error GoTo 0
                endpoints.Add Array(sStart, sEnd)
                entityTypes.Add "Spline"
        End Select
    Next i
    
    If entities.Count = 0 Then
        MsgBox "No lines, arcs, or curves found in the drawing.", vbInformation
        Exit Sub
    End If
    
    ' Ask user if they want to decide on each gap individually or fix all automatically
    Dim fixModeResponse As VbMsgBoxResult
    fixModeResponse = MsgBox("How would you like to handle gaps?" & vbCrLf & vbCrLf & _
                           "Yes - Ask me about each gap individually" & vbCrLf & _
                           "No - Fix all gaps automatically" & vbCrLf & _
                           "Cancel - Exit without fixing", _
                           vbYesNoCancel + vbQuestion, "Gap Fixing Mode")
    
    If fixModeResponse = vbCancel Then Exit Sub
    
    Dim autoFixMode As Boolean
    autoFixMode = (fixModeResponse = vbNo)
    
    ' Search for gaps between endpoints
    Dim gapsFound As Integer
    gapsFound = 0
    
    ' Debug/count info (suppress in auto-fix mode as per new requirement)
    If Not autoFixMode Then
        MsgBox "Found " & entities.Count & " entities to analyze for gaps.", vbInformation
    End If
    
    For i = 1 To entities.Count
        For j = i + 1 To entities.Count
            ' Get endpoints for both entities
            Dim endpoints1 As Variant, endpoints2 As Variant
            endpoints1 = endpoints(i)
            endpoints2 = endpoints(j)
            
            ' Check all combinations of endpoints
            Dim combinations(3) As Variant
            combinations(0) = Array(endpoints1(0), endpoints2(0)) ' Start1 to Start2
            combinations(1) = Array(endpoints1(0), endpoints2(1)) ' Start1 to End2
            combinations(2) = Array(endpoints1(1), endpoints2(0)) ' End1 to Start2
            combinations(3) = Array(endpoints1(1), endpoints2(1)) ' End1 to End2
            
            Dim k As Integer
            For k = 0 To 3
                Dim point1 As Variant, point2 As Variant
                point1 = combinations(k)(0)
                point2 = combinations(k)(1)
                
                ' Calculate distance between points
                Dim distance As Double
                distance = Sqr((point2(0) - point1(0)) ^ 2 + (point2(1) - point1(1)) ^ 2 + (point2(2) - point1(2)) ^ 2)
                
                ' Skip if points are the same (distance = 0) to avoid connecting endpoints of the same entity
                If distance < 0.0000001 Then GoTo NextCombination
                
                ' Check if distance is within gap range
                If distance >= minGap And distance <= maxGap Then
                    gapsFound = gapsFound + 1
                    
                    ' Determine which endpoints are involved
                    Dim endpointDesc As String
                    Select Case k
                        Case 0: endpointDesc = "Start to Start"
                        Case 1: endpointDesc = "Start to End"
                        Case 2: endpointDesc = "End to Start"
                        Case 3: endpointDesc = "End to End"
                    End Select
                    
                    Dim shouldFixGap As Boolean
                    shouldFixGap = False
                    
                    ' Declare displayDistance once for both branches
                    Dim displayDistance As Double
                    displayDistance = distance / unitFactor
                    
                    If autoFixMode Then
                        ' In automatic mode, fix all gaps silently (no per-gap MsgBoxes)
                        shouldFixGap = True
                        ' Optional: still zoom to the gap for visual feedback; comment out next line to disable zoom during auto mode
                        Call ZoomToGap(point1, point2, distance * 10)
                    Else
                        ' In manual mode, ask user about each gap
                        Call ZoomToGap(point1, point2, distance * 10) ' Zoom with 10x buffer
                        
                        ' Ask user if they want to fix this gap
                        userResponse = MsgBox("Gap found between " & entityTypes(i) & " and " & entityTypes(j) & vbCrLf & _
                                            "Connection: " & endpointDesc & vbCrLf & _
                                            "Distance: " & Format(displayDistance, "0.000") & " mm" & vbCrLf & _
                                            "Raw distance: " & Format(distance, "0.000000") & " drawing units" & vbCrLf & _
                                            "Point 1: (" & Format(point1(0), "0.000") & ", " & Format(point1(1), "0.000") & ")" & vbCrLf & _
                                            "Point 2: (" & Format(point2(0), "0.000") & ", " & Format(point2(1), "0.000") & ")" & vbCrLf & _
                                            "Do you want to fix this gap?", vbYesNoCancel + vbQuestion, "Fix Gap?")
                        
                        If userResponse = vbCancel Then
                            Exit Sub
                        ElseIf userResponse = vbYes Then
                            shouldFixGap = True
                        End If
                    End If
                    
                    ' Fix the gap if requested
                    If shouldFixGap Then
                        ' Try to fix by moving endpoints first
                        Dim fixedByMoving As Boolean
                        fixedByMoving = False
                        
                        ' Check if both entities can be modified by moving endpoints
                        If CanMoveEndpoint(entities(i)) And CanMoveEndpoint(entities(j)) Then
                            ' Calculate middle point
                            Dim midPoint(2) As Double
                            midPoint(0) = (point1(0) + point2(0)) / 2
                            midPoint(1) = (point1(1) + point2(1)) / 2
                            midPoint(2) = (point1(2) + point2(2)) / 2
                            
                        ' Try to move endpoints to middle point
                        Dim moved1 As Boolean, moved2 As Boolean
                        Dim newSpline1 As AcadSpline, newSpline2 As AcadSpline
                        Set newSpline1 = Nothing
                        Set newSpline2 = Nothing
                        
                        ' Handle potential arc-to-spline conversion for first entity
                        If entities(i).ObjectName = "AcDbArc" Then
                            moved1 = MoveEndpointWithConversion(entities(i), point1, midPoint, k, True, newSpline1)
                            If moved1 And Not newSpline1 Is Nothing Then
                                ' Note: Collection updates will be handled after both conversions
                                ' to avoid index invalidation during the loop
                            End If
                        Else
                            moved1 = MoveEndpoint(entities(i), point1, midPoint, k, True)
                        End If
                        
                        ' Handle potential arc-to-spline conversion for second entity
                        If entities(j).ObjectName = "AcDbArc" Then
                            moved2 = MoveEndpointWithConversion(entities(j), point2, midPoint, k, False, newSpline2)
                            If moved2 And Not newSpline2 Is Nothing Then
                                ' Note: Collection updates will be handled after both conversions
                                ' to avoid index invalidation during the loop
                            End If
                        Else
                            moved2 = MoveEndpoint(entities(j), point2, midPoint, k, False)
                        End If
                        
                        ' Now safely update collections after both conversions are complete
                        ' Handle first entity replacement (do j first since j > i, to preserve i's index)
                        If moved2 And Not newSpline2 Is Nothing Then
                            On Error Resume Next
                            Dim spline2Start As Variant, spline2End As Variant
                            spline2Start = newSpline2.StartPoint
                            spline2End = newSpline2.EndPoint
                            On Error GoTo 0
                            
                            If Not IsEmpty(spline2Start) And Not IsEmpty(spline2End) Then
                                entities.Remove j
                                entityTypes.Remove j
                                endpoints.Remove j
                                ' Add back at the same logical position, but safely handle bounds
                                If j <= entities.Count Then
                                    entities.Add newSpline2, , j
                                    entityTypes.Add "Spline", , j
                                    endpoints.Add Array(spline2Start, spline2End), , j
                                Else
                                    ' Add at end if original position is now out of bounds
                                    entities.Add newSpline2
                                    entityTypes.Add "Spline"
                                    endpoints.Add Array(spline2Start, spline2End)
                                End If
                            End If
                        End If
                        
                        ' Handle second entity replacement (i is still valid since we processed j first)
                        If moved1 And Not newSpline1 Is Nothing Then
                            On Error Resume Next
                            Dim spline1Start As Variant, spline1End As Variant
                            spline1Start = newSpline1.StartPoint
                            spline1End = newSpline1.EndPoint
                            On Error GoTo 0
                            
                            If Not IsEmpty(spline1Start) And Not IsEmpty(spline1End) Then
                                entities.Remove i
                                entityTypes.Remove i
                                endpoints.Remove i
                                ' Add back at the same logical position, but safely handle bounds
                                If i <= entities.Count Then
                                    entities.Add newSpline1, , i
                                    entityTypes.Add "Spline", , i
                                    endpoints.Add Array(spline1Start, spline1End), , i
                                Else
                                    ' Add at end if original position is now out of bounds
                                    entities.Add newSpline1
                                    entityTypes.Add "Spline"
                                    endpoints.Add Array(spline1Start, spline1End)
                                End If
                            End If
                        End If
                        
                        fixedByMoving = moved1 And moved2
                        End If
                        
                        ' If couldn't fix by moving endpoints, create a connecting line
                        If Not fixedByMoving Then
                            Dim connectingLine As AcadLine
                            Set connectingLine = modelSpace.AddLine(point1, point2)
                            connectingLine.Color = acRed ' Make it red to highlight the fix
                            connectingLine.Linetype = "CONTINUOUS" ' Ensure it's visible
                            connectingLinesCount = connectingLinesCount + 1
                            
                            If Not autoFixMode Then
                                MsgBox "Gap filled with connecting line (entities too complex to modify)!" & vbCrLf & _
                                       "Line color: Red for easy identification", vbInformation
                            End If
                        Else
                            If Not autoFixMode Then
                                MsgBox "Gap fixed by moving endpoints!", vbInformation
                            End If
                        End If
                        
                        ' Regenerate the drawing
                        doc.Regen acActiveViewport
                    End If
                End If
NextCombination:
            Next k
        Next j
    Next i
    
    ' Zoom extents when done
    doc.Application.ZoomExtents
    
    If gapsFound = 0 Then
        MsgBox "No gaps found within the specified range (" & minGapMm & " to " & maxGapMm & " mm).", vbInformation
    Else
        Dim fixSummary As String
        fixSummary = "Gap analysis complete. " & gapsFound & " gaps were found"
        If connectingLinesCount > 0 Then
            fixSummary = fixSummary & vbCrLf & connectingLinesCount & " gaps were filled with red connecting lines"
            fixSummary = fixSummary & vbCrLf & (gapsFound - connectingLinesCount) & " gaps were fixed by moving endpoints"
        Else
            fixSummary = fixSummary & vbCrLf & "All gaps were fixed by moving endpoints"
        End If
        
        If autoFixMode Then
            fixSummary = fixSummary & " automatically."
        Else
            fixSummary = fixSummary & "."
        End If
        
        MsgBox fixSummary, vbInformation
    End If
    
End Sub

Private Function CanMoveEndpoint(entity As AcadEntity) As Boolean
    ' Check if an entity's endpoint can be safely moved
    
    Select Case entity.ObjectName
        Case "AcDbLine"
            CanMoveEndpoint = True ' Lines can always be moved
        Case "AcDbArc"
            CanMoveEndpoint = True ' Arcs can be modified (though complex, will check angle change in ModifyArcEndpoint)
        Case "AcDbPolyline"
            CanMoveEndpoint = True ' Polylines can be modified
        Case "AcDbSpline"
            CanMoveEndpoint = False ' Splines are too complex
        Case "AcDbCircle"
            CanMoveEndpoint = False ' Circles don't have endpoints
        Case "AcDbEllipse"
            CanMoveEndpoint = False ' Ellipses are complex
        Case Else
            CanMoveEndpoint = False ' Unknown entities default to false for safety
    End Select
End Function

Private Sub ZoomToGap(point1 As Variant, point2 As Variant, bufferDistance As Double)
    ' Zoom to the gap region with a buffer
    Dim doc As AcadDocument
    Set doc = ThisDrawing
    
    ' Calculate zoom window
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double
    
    minX = IIf(point1(0) < point2(0), point1(0), point2(0)) - bufferDistance
    maxX = IIf(point1(0) > point2(0), point1(0), point2(0)) + bufferDistance
    minY = IIf(point1(1) < point2(1), point1(1), point2(1)) - bufferDistance
    maxY = IIf(point1(1) > point2(1), point1(1), point2(1)) + bufferDistance
    
    ' Create zoom window points
    Dim lowerLeft(2) As Double, upperRight(2) As Double
    lowerLeft(0) = minX: lowerLeft(1) = minY: lowerLeft(2) = 0
    upperRight(0) = maxX: upperRight(1) = maxY: upperRight(2) = 0
    
    ' Zoom to window
    doc.Application.ZoomWindow lowerLeft, upperRight
End Sub

Private Function MoveEndpointWithConversion(entity As AcadEntity, oldPoint As Variant, newPoint As Variant, combinationIndex As Integer, isFirstEntity As Boolean, ByRef newSpline As AcadSpline) As Boolean
    ' Move endpoint with potential arc-to-spline conversion
    ' Returns the new spline via ByRef if conversion occurs
    
    On Error GoTo ErrorHandler
    
    If entity.ObjectName = "AcDbArc" Then
        Dim arc As AcadArc
        Set arc = entity
        MoveEndpointWithConversion = ModifyArcEndpoint(arc, oldPoint, newPoint, combinationIndex, isFirstEntity, newSpline)
    Else
        ' For non-arc entities, use regular move endpoint
        MoveEndpointWithConversion = MoveEndpoint(entity, oldPoint, newPoint, combinationIndex, isFirstEntity)
    End If
    
    Exit Function
    
ErrorHandler:
    MoveEndpointWithConversion = False
End Function

Private Function MoveEndpoint(entity As AcadEntity, oldPoint As Variant, newPoint As Variant, combinationIndex As Integer, isFirstEntity As Boolean) As Boolean
    ' Move the endpoint of an entity to a new position
    ' Returns True if successful, False if failed
    
    On Error GoTo ErrorHandler
    
    Select Case entity.ObjectName
        Case "AcDbLine"
            Dim line As AcadLine
            Set line = entity
            
            ' Determine which endpoint to move by checking which is closer to oldPoint
            Dim distToStart As Double, distToEnd As Double
            distToStart = Sqr((oldPoint(0) - line.StartPoint(0)) ^ 2 + (oldPoint(1) - line.StartPoint(1)) ^ 2)
            distToEnd = Sqr((oldPoint(0) - line.EndPoint(0)) ^ 2 + (oldPoint(1) - line.EndPoint(1)) ^ 2)
            
            If distToStart < distToEnd Then
                line.StartPoint = newPoint
            Else
                line.EndPoint = newPoint
            End If
            MoveEndpoint = True
            
        Case "AcDbArc"
            ' Arcs cannot be modified directly in this function
            ' Return False to use connecting line instead
            MoveEndpoint = False
            
        Case "AcDbPolyline"
            Dim pline As AcadLWPolyline
            Set pline = entity
            
            ' Get current start and end points
            Dim coords As Variant
            coords = pline.Coordinates
            Dim startPt(2) As Double, endPt(2) As Double
            startPt(0) = coords(0): startPt(1) = coords(1): startPt(2) = 0
            Dim lastIndex As Integer
            lastIndex = UBound(coords) - 1
            endPt(0) = coords(lastIndex - 1): endPt(1) = coords(lastIndex): endPt(2) = 0
            
            ' Determine which endpoint to move
            Dim distToStartPL As Double, distToEndPL As Double
            distToStartPL = Sqr((oldPoint(0) - startPt(0)) ^ 2 + (oldPoint(1) - startPt(1)) ^ 2)
            distToEndPL = Sqr((oldPoint(0) - endPt(0)) ^ 2 + (oldPoint(1) - endPt(1)) ^ 2)
            
            If distToStartPL < distToEndPL Then
                ' Move start point
                Dim startCoord As Variant
                startCoord = pline.Coordinate(0)
                startCoord(0) = newPoint(0)
                startCoord(1) = newPoint(1)
                pline.Coordinate(0) = startCoord
            Else
                ' Move end point
                Dim endCoord As Variant
                Dim lastIdx As Integer
                lastIdx = pline.NumberOfVertices - 1
                endCoord = pline.Coordinate(lastIdx)
                endCoord(0) = newPoint(0)
                endCoord(1) = newPoint(1)
                pline.Coordinate(lastIdx) = endCoord
            End If
            MoveEndpoint = True
            
        Case "AcDbSpline"
            ' Splines are too complex to modify safely
            MoveEndpoint = False
            
        Case Else
            ' Unknown entity type
            MoveEndpoint = False
            
    End Select
    Exit Function
    
ErrorHandler:
    MoveEndpoint = False
End Function

Private Function ModifyArcEndpoint(arc As AcadArc, oldPoint As Variant, newPoint As Variant, combinationIndex As Integer, isFirstEntity As Boolean, ByRef newSpline As AcadSpline) As Boolean
    ' Convert arc to spline and modify endpoint for more reliable gap fixing
    ' This approach provides better flexibility than trying to modify arc geometry
    ' Returns the new spline via ByRef parameter if conversion occurs
    
    On Error GoTo ErrorHandler
    
    ' Calculate gap size to determine if we should modify the arc
    Dim gapSize As Double
    gapSize = Sqr((newPoint(0) - oldPoint(0)) ^ 2 + (newPoint(1) - oldPoint(1)) ^ 2)
    
    ' Only modify arc for reasonable gaps (less than 50% of radius)
    If gapSize >= arc.radius * 0.5 Then
        ' Large gap: don't modify arc, use connecting line instead
        ModifyArcEndpoint = False
        Exit Function
    End If
    
    ' Convert arc to spline and modify the endpoint
    ModifyArcEndpoint = ConvertArcToSplineAndModify(arc, oldPoint, newPoint, newSpline)
    
    Exit Function
    
ErrorHandler:
    ModifyArcEndpoint = False
End Function

Private Function ConvertArcToSplineAndModify(arc As AcadArc, oldPoint As Variant, newPoint As Variant, ByRef newSpline As AcadSpline) As Boolean
    ' Convert arc to spline, modify endpoint, and replace the original arc
    ' Returns the new spline via ByRef parameter
    
    On Error GoTo ErrorHandler
    
    Dim doc As AcadDocument
    Dim modelSpace As AcadModelSpace
    Set doc = ThisDrawing
    Set modelSpace = doc.ModelSpace
    
    ' Get arc properties
    Dim center As Variant, radius As Double
    Dim startAngle As Double, endAngle As Double
    center = arc.center
    radius = arc.radius
    startAngle = arc.StartAngle
    endAngle = arc.EndAngle
    
    ' Calculate current endpoints
    Dim currentStartPt(2) As Double, currentEndPt(2) As Double
    currentStartPt(0) = center(0) + radius * Cos(startAngle)
    currentStartPt(1) = center(1) + radius * Sin(startAngle)
    currentStartPt(2) = center(2)
    currentEndPt(0) = center(0) + radius * Cos(endAngle)
    currentEndPt(1) = center(1) + radius * Sin(endAngle)
    currentEndPt(2) = center(2)
    
    ' Determine which endpoint to modify
    Dim distToStart As Double, distToEnd As Double
    distToStart = Sqr((oldPoint(0) - currentStartPt(0)) ^ 2 + (oldPoint(1) - currentStartPt(1)) ^ 2)
    distToEnd = Sqr((oldPoint(0) - currentEndPt(0)) ^ 2 + (oldPoint(1) - currentEndPt(1)) ^ 2)
    
    ' Generate control points for the spline based on the arc
    Dim numPoints As Integer
    numPoints = 7 ' Use 7 points for smooth curve approximation
    
    Dim splinePoints() As Double
    ReDim splinePoints((numPoints * 3) - 1) ' 3 coordinates per point
    
    ' Calculate the angular span of the arc
    Dim totalAngle As Double
    totalAngle = endAngle - startAngle
    
    ' Handle angle wrapping for arcs that cross 0 degrees
    If totalAngle < 0 Then
        totalAngle = totalAngle + 2 * 3.14159265358979 ' Add 2*PI
    End If
    
    ' Generate points along the arc curve
    Dim i As Integer
    For i = 0 To numPoints - 1
        Dim currentAngle As Double
        If i = 0 Then
            ' First point - use modified start or original start
            If distToStart < distToEnd Then
                ' Modifying start point
                splinePoints(i * 3) = newPoint(0)
                splinePoints(i * 3 + 1) = newPoint(1)
                splinePoints(i * 3 + 2) = newPoint(2)
            Else
                ' Original start point
                splinePoints(i * 3) = currentStartPt(0)
                splinePoints(i * 3 + 1) = currentStartPt(1)
                splinePoints(i * 3 + 2) = currentStartPt(2)
            End If
        ElseIf i = numPoints - 1 Then
            ' Last point - use modified end or original end
            If distToStart >= distToEnd Then
                ' Modifying end point
                splinePoints(i * 3) = newPoint(0)
                splinePoints(i * 3 + 1) = newPoint(1)
                splinePoints(i * 3 + 2) = newPoint(2)
            Else
                ' Original end point
                splinePoints(i * 3) = currentEndPt(0)
                splinePoints(i * 3 + 1) = currentEndPt(1)
                splinePoints(i * 3 + 2) = currentEndPt(2)
            End If
        Else
            ' Intermediate points - calculate along arc
            Dim t As Double
            t = CDbl(i) / CDbl(numPoints - 1)
            currentAngle = startAngle + t * totalAngle
            
            splinePoints(i * 3) = center(0) + radius * Cos(currentAngle)
            splinePoints(i * 3 + 1) = center(1) + radius * Sin(currentAngle)
            splinePoints(i * 3 + 2) = center(2)
        End If
    Next i
    
    ' Create the spline
    Set newSpline = modelSpace.AddSpline(splinePoints, Empty, Empty)
    
    ' Copy properties from original arc
    newSpline.Color = arc.Color
    newSpline.Layer = arc.Layer
    newSpline.Linetype = arc.Linetype
    newSpline.LinetypeScale = arc.LinetypeScale
    newSpline.Lineweight = arc.Lineweight
    
    ' Delete the original arc
    arc.Delete
    
    ConvertArcToSplineAndModify = True
    Exit Function
    
ErrorHandler:
    ConvertArcToSplineAndModify = False
End Function

Private Sub UpdateEntityEndpoints(entity As AcadEntity, endpoints As Collection, entityIndex As Integer)
    ' Update the endpoint information in the collection after entity modification
    ' This is crucial for arcs whose endpoints may change after center/radius adjustments
    ' Instead of trying to modify collection items directly, we'll recreate the endpoints array
    
    Dim updatedEndpoints As Variant
    
    Select Case entity.ObjectName
        Case "AcDbLine"
            Dim line As AcadLine
            Set line = entity
            updatedEndpoints = Array(line.StartPoint, line.EndPoint)
            
        Case "AcDbArc"
            Dim arc As AcadArc
            Set arc = entity
            ' Recalculate arc endpoints with current center, radius, and angles
            Dim arcStartPt(2) As Double, arcEndPt(2) As Double
            arcStartPt(0) = arc.center(0) + arc.radius * Cos(arc.StartAngle)
            arcStartPt(1) = arc.center(1) + arc.radius * Sin(arc.StartAngle)
            arcStartPt(2) = arc.center(2)
            arcEndPt(0) = arc.center(0) + arc.radius * Cos(arc.EndAngle)
            arcEndPt(1) = arc.center(1) + arc.radius * Sin(arc.EndAngle)
            arcEndPt(2) = arc.center(2)
            updatedEndpoints = Array(arcStartPt, arcEndPt)
            
        Case "AcDbPolyline"
            Dim pline As AcadLWPolyline
            Set pline = entity
            ' Get updated coordinates
            Dim coords As Variant
            coords = pline.Coordinates
            Dim startPt(2) As Double, endPt(2) As Double
            startPt(0) = coords(0): startPt(1) = coords(1): startPt(2) = 0
            Dim lastIndex As Integer
            lastIndex = UBound(coords) - 1
            endPt(0) = coords(lastIndex - 1): endPt(1) = coords(lastIndex): endPt(2) = 0
            updatedEndpoints = Array(startPt, endPt)
            
        Case "AcDbSpline"
            Dim spline As AcadSpline
            Set spline = entity
            ' Safely attempt to read start/end points; fall back to bounding box if unavailable
            On Error Resume Next
            Dim tStart As Variant, tEnd As Variant
            tStart = spline.StartPoint
            tEnd = spline.EndPoint
            If Err.Number <> 0 Then
                Err.Clear
                Dim bbMin2(2) As Double, bbMax2(2) As Double
                spline.GetBoundingBox bbMin2, bbMax2
                tStart = bbMin2
                tEnd = bbMax2
            End If
            On Error GoTo 0
            updatedEndpoints = Array(tStart, tEnd)
    End Select
    
    ' Replace the item in the collection by removing and adding at the same position
    ' First, check if we need to preserve order
    If entityIndex <= endpoints.Count Then
        endpoints.Remove entityIndex
        If entityIndex = 1 Then
            endpoints.Add updatedEndpoints, , 1
        ElseIf entityIndex > endpoints.Count Then
            endpoints.Add updatedEndpoints
        Else
            endpoints.Add updatedEndpoints, , entityIndex
        End If
    End If
End Sub

' VBA does not have Atan2, so we define our own
Private Function Atan2(y As Double, x As Double) As Double
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            Atan2 = Atn(y / x) + 3.14159265358979
        Else
            Atan2 = Atn(y / x) - 3.14159265358979
        End If
    Else ' x = 0
        If y > 0 Then
            Atan2 = 3.14159265358979 / 2
        ElseIf y < 0 Then
            Atan2 = -3.14159265358979 / 2
        Else
            Atan2 = 0 ' undefined, return 0
        End If
    End If
End Function