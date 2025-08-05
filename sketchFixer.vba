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
    response = InputBox("Enter minimum gap value in mm:", "Minimum Gap", "0.001")
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
                endpoints.Add Array(spline.StartPoint, spline.EndPoint)
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
    
    ' Debug: Show how many entities were found
    MsgBox "Found " & entities.Count & " entities to analyze for gaps.", vbInformation
    
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
                        ' In automatic mode, fix all gaps
                        shouldFixGap = True
                        ' Still zoom to show the user what's being fixed
                        Call ZoomToGap(point1, point2, distance * 10)
                        ' Brief message about what's being fixed
                        MsgBox "Fixing gap between " & entityTypes(i) & " and " & entityTypes(j) & vbCrLf & _
                               "Connection: " & endpointDesc & vbCrLf & _
                               "Distance: " & Format(displayDistance, "0.000") & " mm", _
                               vbInformation, "Auto-Fixing Gap " & gapsFound
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
                            moved1 = MoveEndpoint(entities(i), point1, midPoint, k, True)
                            moved2 = MoveEndpoint(entities(j), point2, midPoint, k, False)
                            
                            fixedByMoving = moved1 And moved2
                            
                            ' If both movements were successful, update the endpoints collections
                            ' This is important for arcs whose endpoints may have changed position
                            If fixedByMoving Then
                                Call UpdateEntityEndpoints(entities(i), endpoints, i)
                                Call UpdateEntityEndpoints(entities(j), endpoints, j)
                            End If
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
            Dim arc As AcadArc
            Set arc = entity
            
            ' For arcs, use the specialized function
            MoveEndpoint = ModifyArcEndpoint(arc, oldPoint, newPoint, combinationIndex, isFirstEntity)
            
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

Private Function ModifyArcEndpoint(arc As AcadArc, oldPoint As Variant, newPoint As Variant, combinationIndex As Integer, isFirstEntity As Boolean) As Boolean
    ' Modify arc endpoint by adjusting center and/or radius to place endpoint at exact target position
    ' Returns True if successful, False if failed
    
    On Error GoTo ErrorHandler
    
    Dim startAngle As Double, endAngle As Double
    Dim center As Variant, radius As Double
    
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
    
    ' Calculate distances to determine which endpoint was the old point
    Dim distToStart As Double, distToEnd As Double
    distToStart = Sqr((oldPoint(0) - currentStartPt(0)) ^ 2 + (oldPoint(1) - currentStartPt(1)) ^ 2)
    distToEnd = Sqr((oldPoint(0) - currentEndPt(0)) ^ 2 + (oldPoint(1) - currentEndPt(1)) ^ 2)
    
    ' Calculate distance from new point to current arc center
    Dim distToCenter As Double
    distToCenter = Sqr((newPoint(0) - center(0)) ^ 2 + (newPoint(1) - center(1)) ^ 2)
    
    ' Check if the new point is at a reasonable distance from center
    ' If too far from the original radius, adjust approach
    Dim radiusChange As Double
    radiusChange = Abs(distToCenter - radius)
    
    If distToStart < distToEnd Then
        ' Modifying start point
        If radiusChange / radius < 0.1 Then ' Less than 10% radius change
            ' Method 1: Adjust both radius and angle to place start point exactly at new position
            Dim newRadius As Double
            newRadius = distToCenter
            Dim newStartAngle As Double
            newStartAngle = Atan2(newPoint(1) - center(1), newPoint(0) - center(0))
            
            ' Update arc properties
            arc.radius = newRadius
            arc.StartAngle = newStartAngle
            
            ModifyArcEndpoint = True
        Else
            ' Method 2: Move center to maintain radius and place endpoint exactly
            Dim newCenter(2) As Double
            Dim directionAngle As Double
            directionAngle = Atan2(newPoint(1) - center(1), newPoint(0) - center(0))
            
            ' Calculate new center position to place start point exactly at new position
            newCenter(0) = newPoint(0) - radius * Cos(startAngle)
            newCenter(1) = newPoint(1) - radius * Sin(startAngle)
            newCenter(2) = center(2)
            
            arc.center = newCenter
            ModifyArcEndpoint = True
        End If
    Else
        ' Modifying end point
        If radiusChange / radius < 0.1 Then ' Less than 10% radius change
            ' Method 1: Adjust both radius and angle to place end point exactly at new position
            Dim newRadiusEnd As Double
            newRadiusEnd = distToCenter
            Dim newEndAngle As Double
            newEndAngle = Atan2(newPoint(1) - center(1), newPoint(0) - center(0))
            
            ' Update arc properties
            arc.radius = newRadiusEnd
            arc.EndAngle = newEndAngle
            
            ModifyArcEndpoint = True
        Else
            ' Method 2: Move center to maintain radius and place endpoint exactly
            Dim newCenterEnd(2) As Double
            
            ' Calculate new center position to place end point exactly at new position
            newCenterEnd(0) = newPoint(0) - radius * Cos(endAngle)
            newCenterEnd(1) = newPoint(1) - radius * Sin(endAngle)
            newCenterEnd(2) = center(2)
            
            arc.center = newCenterEnd
            ModifyArcEndpoint = True
        End If
    End If
    Exit Function
    
ErrorHandler:
    ModifyArcEndpoint = False
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
            updatedEndpoints = Array(spline.StartPoint, spline.EndPoint)
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