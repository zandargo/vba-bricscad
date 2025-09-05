' BricsCAD VBA Macro: Polygon Detection
' Detects polygons formed by connected lines in the drawing
' Finds chains of connected lines and reports polygon information

Option Explicit

' Global tolerance value for line connection detection
Private Const GLOBAL_TOLERANCE As Double = 0.01

' List of lateral widths to check for hexagons (in drawing units)
Private Const LATERAL_WIDTHS As String = "5,10"  ' Example: "5,10,15,25"
Private Const WIDTH_TOLERANCE_PERCENT As Double = 0.01 ' 1% margin of error

' Progress bar settings
Private Const PROGRESS_THRESHOLD As Integer = 100 ' Show progress bar if more than this many lines
Private Const PROGRESS_INTERVAL As Integer = 10 ' Update progress every N percent

Public Sub DetectPolygons()
    Dim doc As AcadDocument
    Set doc = ThisDrawing
    
    ' Ensure Hexagonos layer exists and is green
    Call EnsureHexagonosLayer
    
    ' Ensure Puncionadeira layer exists and is red
    Call EnsurePuncionadeiraLayer
    
    Dim lines As Collection
    Set lines = New Collection
    
    Dim ent As AcadEntity
    ' Collect all lines in the drawing
    For Each ent In doc.ModelSpace
        If ent.ObjectName = "AcDbLine" Then
            lines.Add ent
        End If
    Next
    
    Debug.Print "Found " & lines.Count & " lines in the drawing"
    Debug.Print "Starting polygon detection with tolerance: " & GLOBAL_TOLERANCE
    Debug.Print "========================================================"
    
    ' Determine if we need to show progress
    Dim showProgress As Boolean
    showProgress = lines.Count > PROGRESS_THRESHOLD
    
    If showProgress Then
        Debug.Print "Processing " & lines.Count & " lines - progress will be shown..."
        Debug.Print "Progress: [          ] 0%"
    End If
    
    Dim processedLines As Collection
    Set processedLines = New Collection
    
    Dim polygonCount As Long
    Dim hexagonCount As Long
    Dim puncionadeiraCount As Long
    polygonCount = 0
    hexagonCount = 0
    puncionadeiraCount = 0
    
    Dim lineIndex As Long
    Dim startLine As AcadLine
    Dim polygon As Collection
    Dim polyLine As AcadLine
    Dim j As Long
    
    ' Progress tracking variables
    Dim lastProgressPercent As Long
    lastProgressPercent = 0

    For lineIndex = 1 To lines.Count
        ' Update progress bar if needed
        If showProgress Then
            Dim currentPercent As Long
            currentPercent = CLng((lineIndex * 100) / IIf(lines.Count = 0, 1, lines.Count))
            
            If currentPercent >= lastProgressPercent + PROGRESS_INTERVAL Then
                Call UpdateProgressBar(currentPercent)
                lastProgressPercent = currentPercent
            End If
        End If
        
        Set startLine = lines(lineIndex)

        ' Skip if this line was already processed in a polygon
        If Not IsLineProcessed(startLine, processedLines) Then
            Set polygon = FindPolygon(startLine, lines, GLOBAL_TOLERANCE)

            If Not polygon Is Nothing Then
                If polygon.Count >= 3 Then
                    polygonCount = polygonCount + 1
                    Call ReportPolygon(polygon, polygonCount)
                    
                    ' Check if it's a hexagon and move to appropriate layer
                    If polygon.Count = 6 Then
                        hexagonCount = hexagonCount + 1
                        
                        ' Check if hexagon has lateral width matching our list
                        Dim lateralWidth As Double
                        lateralWidth = GetHexagonLateralWidth(polygon)
                        
                        If IsWidthInList(lateralWidth) Then
                            puncionadeiraCount = puncionadeiraCount + 1
                            Call MovePolygonToPuncionadeiraLayer(polygon)
                            Debug.Print "  --> Hexagon with lateral width " & Format(lateralWidth, "0.000") & " moved to 'Puncionadeira' layer"
                        Else
                            Call MovePolygonToHexagonosLayer(polygon)
                            Debug.Print "  --> Hexagon with lateral width " & Format(lateralWidth, "0.000") & " moved to 'Hexagonos' layer"
                        End If
                    End If

                    ' Mark all lines in this polygon as processed
                    For j = 1 To polygon.Count
                        Set polyLine = polygon(j)
                        processedLines.Add polyLine
                    Next j
                End If
            End If
        End If
    Next lineIndex
    
    ' Complete progress bar if shown
    If showProgress Then
        Call UpdateProgressBar(100)
        Debug.Print "" ' Add blank line after progress
    End If
    
    Debug.Print ""
    Debug.Print "Total polygons found: " & polygonCount
    Debug.Print "Total hexagons found: " & hexagonCount
    Debug.Print "Hexagons moved to 'Hexagonos' layer: " & (hexagonCount - puncionadeiraCount)
    Debug.Print "Hexagons moved to 'Puncionadeira' layer: " & puncionadeiraCount
    Debug.Print "Polygon detection completed."
End Sub

' Finds a polygon starting from a given line
Private Function FindPolygon(startLine As AcadLine, allLines As Collection, tolerance As Double) As Collection
    Dim polygon As Collection
    Set polygon = New Collection
    
    Dim visitedLines As Collection
    Set visitedLines = New Collection
    
    polygon.Add startLine
    visitedLines.Add startLine
    
    Dim currentLine As AcadLine
    Set currentLine = startLine
    Dim currentEndPoint As Variant
    currentEndPoint = currentLine.EndPoint
    
    Dim maxIterations As Long
    maxIterations = 50 ' Prevent infinite loops
    Dim iteration As Long
    iteration = 0
    
    Do While iteration < maxIterations
        iteration = iteration + 1
        
        Dim nextLine As AcadLine
        Set nextLine = FindConnectedLine(currentEndPoint, allLines, visitedLines, tolerance)
        
        If nextLine Is Nothing Then
            ' No more connected lines found
            Exit Do
        End If
        
        ' Determine which end of the next line to use
        Dim nextEndPoint As Variant
        If IsPointsEqual(nextLine.StartPoint, currentEndPoint, tolerance) Then
            nextEndPoint = nextLine.EndPoint
        Else
            nextEndPoint = nextLine.StartPoint
        End If
        
        ' Check if this line connects back to the start (closes the polygon)
        If IsPointsEqual(nextEndPoint, startLine.StartPoint, tolerance) Then
            polygon.Add nextLine
            Exit Do ' Closed polygon found
        End If
        
        ' Add the line to the polygon and continue
        polygon.Add nextLine
        visitedLines.Add nextLine
        Set currentLine = nextLine
        currentEndPoint = nextEndPoint
    Loop
    
    ' Verify it's a closed polygon
    If polygon.Count >= 3 Then
        Dim firstLine As AcadLine
        Dim lastLine As AcadLine
        Set firstLine = polygon(1)
        Set lastLine = polygon(polygon.Count)
        
        ' Check if polygon is properly closed
        Dim lastEndPoint As Variant
        If IsPointsEqual(lastLine.StartPoint, currentEndPoint, tolerance) Then
            lastEndPoint = lastLine.EndPoint
        Else
            lastEndPoint = lastLine.StartPoint
        End If
        
        If Not IsPointsEqual(lastEndPoint, firstLine.StartPoint, tolerance) Then
            ' Not a closed polygon, return empty collection
            Set polygon = New Collection
        End If
    Else
        ' Less than 3 sides, not a polygon
        Set polygon = New Collection
    End If
    
    Set FindPolygon = polygon
End Function

' Finds a line connected to the given point
Private Function FindConnectedLine(point As Variant, allLines As Collection, visitedLines As Collection, tolerance As Double) As AcadLine
    Dim line As AcadLine
        Dim i As Long
    
    For i = 1 To allLines.Count
        Set line = allLines(i)
        
        ' Skip if already visited
        If Not IsLineInCollection(line, visitedLines) Then
            ' Check if the line's start point connects to our point
            If IsPointsEqual(line.StartPoint, point, tolerance) Then
                Set FindConnectedLine = line
                Exit Function
            End If
            ' Also check if the line's end point connects to our point
            If IsPointsEqual(line.EndPoint, point, tolerance) Then
                Set FindConnectedLine = line
                Exit Function
            End If
        End If
    Next i
    
    Set FindConnectedLine = Nothing
End Function

' Checks if two points are equal within tolerance
Private Function IsPointsEqual(point1 As Variant, point2 As Variant, tolerance As Double) As Boolean
    Dim dx As Double, dy As Double, dz As Double
    dx = point1(0) - point2(0)
    dy = point1(1) - point2(1)
    dz = point1(2) - point2(2)
    
    IsPointsEqual = Sqr(dx * dx + dy * dy + dz * dz) <= tolerance
End Function

' Checks if a line is in a collection
Private Function IsLineInCollection(line As AcadLine, collection As Collection) As Boolean
    Dim item As AcadLine
        Dim i As Long
    
    For i = 1 To collection.Count
        Set item = collection(i)
        If item.Handle = line.Handle Then
            IsLineInCollection = True
            Exit Function
        End If
    Next i
    
    IsLineInCollection = False
End Function

' Checks if a line was already processed
Private Function IsLineProcessed(line As AcadLine, processedLines As Collection) As Boolean
    IsLineProcessed = IsLineInCollection(line, processedLines)
End Function

' Reports information about a found polygon
Private Sub ReportPolygon(polygon As Collection, polygonNumber As Long)
    Debug.Print "Polygon #" & polygonNumber & ":"
    Debug.Print "  Number of sides: " & polygon.Count
    Debug.Print "  Line handles: ";
    
    Dim line As AcadLine
    Dim i As Long
    For i = 1 To polygon.Count
        Set line = polygon(i)
        Debug.Print line.Handle;
        If i < polygon.Count Then Debug.Print ", ";
    Next i
    Debug.Print ""
    
    ' Calculate and report polygon properties
    Dim perimeter As Double
    perimeter = CalculatePolygonPerimeter(polygon)
    Debug.Print "  Perimeter: " & Format(perimeter, "0.000")
    
    Dim area As Double
    area = CalculatePolygonArea(polygon)
    Debug.Print "  Area: " & Format(area, "0.000")
    
    Debug.Print ""
End Sub

' Calculates the perimeter of a polygon
Private Function CalculatePolygonPerimeter(polygon As Collection) As Double
    Dim totalLength As Double
    totalLength = 0
    
    Dim line As AcadLine
        Dim i As Long
    For i = 1 To polygon.Count
        Set line = polygon(i)
        totalLength = totalLength + line.Length
    Next i
    
    CalculatePolygonPerimeter = totalLength
End Function

' Calculates the area of a polygon using the shoelace formula
Private Function CalculatePolygonArea(polygon As Collection) As Double
    If polygon.Count < 3 Then
        CalculatePolygonArea = 0
        Exit Function
    End If
    
    Dim area As Double
    area = 0
    
    Dim line As AcadLine
    Dim i As Long
    For i = 1 To polygon.Count
        Set line = polygon(i)
        Dim x1 As Double, y1 As Double
        Dim x2 As Double, y2 As Double
        
        x1 = line.StartPoint(0)
        y1 = line.StartPoint(1)
        x2 = line.EndPoint(0)
        y2 = line.EndPoint(1)
        
        area = area + (x1 * y2 - x2 * y1)
    Next i
    
    CalculatePolygonArea = Abs(area) / 2
End Function

' Helper function to highlight found polygons (optional)
Public Sub HighlightPolygon(polygon As Collection, colorIndex As Integer)
    Dim line As AcadLine
    Dim i As Long
    For i = 1 To polygon.Count
        Set line = polygon(i)
        line.color = colorIndex
    Next i
    
    ThisDrawing.Regen acAllViewports
End Sub

' Ensures the Hexagonos layer exists and is configured as green
Private Sub EnsureHexagonosLayer()
    Dim doc As AcadDocument
    Set doc = ThisDrawing
    
    Dim layerName As String
    layerName = "Hexagonos"
    
    Dim hexagonosLayer As AcadLayer
    
    ' Try to get the layer, create it if it doesn't exist
    On Error GoTo CreateLayer
    Set hexagonosLayer = doc.Layers(layerName)
    GoTo ConfigureLayer
    
CreateLayer:
    Set hexagonosLayer = doc.Layers.Add(layerName)
    Debug.Print "Created layer '" & layerName & "'"
    
ConfigureLayer:
    ' Set layer color to green (color index 3)
   '  hexagonosLayer.color = acGreen
    hexagonosLayer.Linetype = "Continuous"
    
    Debug.Print "Configured layer '" & layerName & "' as green"
End Sub

' Moves all lines of a polygon to the Hexagonos layer
Private Sub MovePolygonToHexagonosLayer(polygon As Collection)
    Dim line As AcadLine
        Dim i As Long
    
    For i = 1 To polygon.Count
        Set line = polygon(i)
        ' Ensure the line uses the layer color by setting its color to ByLayer
        line.Color = acByLayer
        line.Layer = "Hexagonos"
    Next i
End Sub

' Ensures the Puncionadeira layer exists and is configured as red
Private Sub EnsurePuncionadeiraLayer()
    Dim doc As AcadDocument
    Set doc = ThisDrawing
    
    Dim layerName As String
    layerName = "Puncionadeira"
    
    Dim puncionadeiraLayer As AcadLayer
    
    ' Try to get the layer, create it if it doesn't exist
    On Error GoTo CreateLayer
    Set puncionadeiraLayer = doc.Layers(layerName)
    GoTo ConfigureLayer
    
CreateLayer:
    Set puncionadeiraLayer = doc.Layers.Add(layerName)
    Debug.Print "Created layer '" & layerName & "'"
    
ConfigureLayer:
    ' Set layer color to red (color index 1)
    puncionadeiraLayer.color = acRed
    puncionadeiraLayer.Linetype = "Continuous"
    
    Debug.Print "Configured layer '" & layerName & "' as red"
End Sub

' Moves all lines of a polygon to the Puncionadeira layer
Private Sub MovePolygonToPuncionadeiraLayer(polygon As Collection)
    Dim line As AcadLine
        Dim i As Long
    
    For i = 1 To polygon.Count
        Set line = polygon(i)
        ' Ensure the line uses the layer color by setting its color to ByLayer
        line.Color = acByLayer
        line.Layer = "Puncionadeira"
    Next i
End Sub

' Calculates the lateral width of a hexagon (distance between parallel sides)
Private Function GetHexagonLateralWidth(polygon As Collection) As Double
    If polygon.Count <> 6 Then
        GetHexagonLateralWidth = 0
        Exit Function
    End If
    
    ' For a regular hexagon, we need to find the distance between parallel sides
    ' We'll calculate the distance from each line to all other lines and find the maximum
    Dim maxDistance As Double
    maxDistance = 0
    
    Dim line1 As AcadLine, line2 As AcadLine
        Dim i As Long, j As Long
    
    For i = 1 To polygon.Count
        Set line1 = polygon(i)
        For j = i + 1 To polygon.Count
            Set line2 = polygon(j)
            
            ' Check if lines are approximately parallel
            If AreLinesParallel(line1, line2) Then
                Dim distance As Double
                distance = DistanceBetweenParallelLines(line1, line2)
                If distance > maxDistance Then
                    maxDistance = distance
                End If
            End If
        Next j
    Next i
    
    GetHexagonLateralWidth = maxDistance
End Function

' Checks if two lines are approximately parallel
Private Function AreLinesParallel(line1 As AcadLine, line2 As AcadLine) As Boolean
    Dim dx1 As Double, dy1 As Double
    Dim dx2 As Double, dy2 As Double
    
    ' Calculate direction vectors
    dx1 = line1.EndPoint(0) - line1.StartPoint(0)
    dy1 = line1.EndPoint(1) - line1.StartPoint(1)
    dx2 = line2.EndPoint(0) - line2.StartPoint(0)
    dy2 = line2.EndPoint(1) - line2.StartPoint(1)
    
    ' Normalize the vectors
    Dim len1 As Double, len2 As Double
    len1 = Sqr(dx1 * dx1 + dy1 * dy1)
    len2 = Sqr(dx2 * dx2 + dy2 * dy2)
    
    If len1 = 0 Or len2 = 0 Then
        AreLinesParallel = False
        Exit Function
    End If
    
    dx1 = dx1 / len1
    dy1 = dy1 / len1
    dx2 = dx2 / len2
    dy2 = dy2 / len2
    
    ' Calculate cross product (for 2D, this is dx1*dy2 - dy1*dx2)
    Dim crossProduct As Double
    crossProduct = Abs(dx1 * dy2 - dy1 * dx2)
    
    ' Lines are parallel if cross product is close to 0
    AreLinesParallel = crossProduct < 0.1 ' Tolerance for "approximately parallel"
End Function

' Calculates distance between two parallel lines
Private Function DistanceBetweenParallelLines(line1 As AcadLine, line2 As AcadLine) As Double
    ' Use point-to-line distance formula
    ' Distance from point (x0,y0) to line ax + by + c = 0 is |ax0 + by0 + c| / sqrt(a^2 + b^2)
    
    ' Get line1 equation coefficients (ax + by + c = 0)
    Dim dx As Double, dy As Double
    dx = line1.EndPoint(0) - line1.StartPoint(0)
    dy = line1.EndPoint(1) - line1.StartPoint(1)
    
    ' Normal vector to line1 is (-dy, dx)
    Dim a As Double, b As Double, c As Double
    a = -dy
    b = dx
    c = dy * line1.StartPoint(0) - dx * line1.StartPoint(1)
    
    ' Calculate distance from line2's start point to line1
    Dim x0 As Double, y0 As Double
    x0 = line2.StartPoint(0)
    y0 = line2.StartPoint(1)
    
    Dim distance As Double
    distance = Abs(a * x0 + b * y0 + c) / Sqr(a * a + b * b)
    
    DistanceBetweenParallelLines = distance
End Function

' Checks if a width matches any value in our lateral widths list (within tolerance)
Private Function IsWidthInList(width As Double) As Boolean
    Dim widthArray As Variant
    widthArray = Split(LATERAL_WIDTHS, ",")
    
    Dim i As Long
        For i = 0 To UBound(widthArray)
        Dim targetWidth As Double
        targetWidth = CDbl(Trim(widthArray(i)))
        
        ' Check if width is within tolerance percentage
        Dim tolerance As Double
        tolerance = targetWidth * WIDTH_TOLERANCE_PERCENT
        
        If Abs(width - targetWidth) <= tolerance Then
            IsWidthInList = True
            Exit Function
        End If
    Next i
    
    IsWidthInList = False
End Function

' Updates and displays a text-based progress bar
Private Sub UpdateProgressBar(percent As Long)
    Dim progressBar As String
    Dim barLength As Long
    barLength = 10
    
    Dim filledLength As Long
    filledLength = CLng((percent * barLength) / 100)
    
    ' Build progress bar string
    progressBar = "Progress: ["
    
    Dim i As Long
    For i = 1 To barLength
        If i <= filledLength Then
            progressBar = progressBar & "="
        Else
            progressBar = progressBar & " "
        End If
    Next i
    
    progressBar = progressBar & "] " & percent & "%"
    
    ' Clear previous line and print new progress
    Debug.Print Chr(13) & progressBar;
    
    ' Force immediate update of the debug window
    DoEvents
End Sub
