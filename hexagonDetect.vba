' BricsCAD VBA Macro: Polygon Detection
' Detects polygons formed by connected lines in the drawing
' Finds chains of connected lines and reports polygon information

Option Explicit

' Global tolerance value for line connection detection
Private Const GLOBAL_TOLERANCE As Double = 0.01

Public Sub DetectPolygons()
    Dim doc As AcadDocument
    Set doc = ThisDrawing
    
    ' Ensure Hexagonos layer exists and is green
    Call EnsureHexagonosLayer
    
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
    
    Dim processedLines As Collection
    Set processedLines = New Collection
    
    Dim polygonCount As Integer
    Dim hexagonCount As Integer
    polygonCount = 0
    hexagonCount = 0
    
    Dim lineIndex As Integer
    Dim startLine As AcadLine
    Dim polygon As Collection
    Dim polyLine As AcadLine
    Dim j As Integer

    For lineIndex = 1 To lines.Count
        Set startLine = lines(lineIndex)

        ' Skip if this line was already processed in a polygon
        If Not IsLineProcessed(startLine, processedLines) Then
            Set polygon = FindPolygon(startLine, lines, GLOBAL_TOLERANCE)

            If Not polygon Is Nothing Then
                If polygon.Count >= 3 Then
                    polygonCount = polygonCount + 1
                    Call ReportPolygon(polygon, polygonCount)
                    
                    ' Check if it's a hexagon and move to Hexagonos layer
                    If polygon.Count = 6 Then
                        hexagonCount = hexagonCount + 1
                        Call MovePolygonToHexagonosLayer(polygon)
                        Debug.Print "  --> Moved hexagon to 'Hexagonos' layer"
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
    
    Debug.Print ""
    Debug.Print "Total polygons found: " & polygonCount
    Debug.Print "Total hexagons found and moved to 'Hexagonos' layer: " & hexagonCount
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
    
    Dim maxIterations As Integer
    maxIterations = 50 ' Prevent infinite loops
    Dim iteration As Integer
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
    Dim i As Integer
    
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
    Dim i As Integer
    
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
Private Sub ReportPolygon(polygon As Collection, polygonNumber As Integer)
    Debug.Print "Polygon #" & polygonNumber & ":"
    Debug.Print "  Number of sides: " & polygon.Count
    Debug.Print "  Line handles: ";
    
    Dim line As AcadLine
    Dim i As Integer
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
    Dim i As Integer
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
    Dim i As Integer
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
    Dim i As Integer
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
    hexagonosLayer.color = acGreen
    hexagonosLayer.Linetype = "Continuous"
    
    Debug.Print "Configured layer '" & layerName & "' as green"
End Sub

' Moves all lines of a polygon to the Hexagonos layer
Private Sub MovePolygonToHexagonosLayer(polygon As Collection)
    Dim line As AcadLine
    Dim i As Integer
    
    For i = 1 To polygon.Count
        Set line = polygon(i)
        line.Layer = "Hexagonos"
    Next i
End Sub
