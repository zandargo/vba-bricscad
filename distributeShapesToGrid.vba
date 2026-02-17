Option Explicit

' Distribute detected shapes onto a grid of cells.
' Steps:
' 1) Hide formPerfisul01 and ask user to window-select the shapes.
' 2) Detect closed regions from the selection (outer shapes only).
' 3) For each region, gather its entities, auto-orient to minimize height, and record width/height.
' 4) Get the maximum region width (maxWidth).
' 5) Ask user to select the grid; detect columns, rows, cell width/height, and centers.
' 6) Scale shapes so the widest shape fits the grid cell width.
' 7) Move each rotated/scaled shape to successive cell centers. If shapes exceed cells, create extra rows beneath the grid.

Private Type Point2D
	x As Double
	y As Double
End Type

Public Sub DistributeShapesToGrid()
	On Error GoTo ErrHandler
	If Not formPerfisul01 Is Nothing Then
		On Error Resume Next
		formPerfisul01.Hide
		On Error GoTo ErrHandler
	End If
    
	Dim doc As AcadDocument
	Set doc = ThisDrawing
	doc.StartUndoMark
	Dim shapesLayer As AcadLayer
	Set shapesLayer = EnsureShapesLayer(doc)
    
	Dim shapeSS As AcadSelectionSet
	Set shapeSS = PrepareSelectionSet(doc, "DSG_SHAPES")
	doc.Utility.Prompt vbCr & "Selecione as formas (window selection)..." & vbCr
	shapeSS.SelectOnScreen
	If shapeSS.Count = 0 Then
		MsgBox "Nenhum objeto selecionado para distribuir.", vbExclamation, "Distribuir Formas"
		GoTo Cleanup
	End If
	NormalizeSelectedLayers shapeSS
    
	Dim allRegions As Collection
	Dim outerRegions As Collection
	Set outerRegions = DetectOuterRegionsFromSelection(doc, shapeSS, allRegions)
	If outerRegions Is Nothing Or outerRegions.Count = 0 Then
		MsgBox "Não foi possível detectar regiões fechadas.", vbExclamation, "Distribuir Formas"
		GoTo Cleanup
	End If
    
	Dim regionEntities As Collection
	Set regionEntities = New Collection
	Dim regionLabels() As String
	Dim regionCenters() As Variant
	Dim regionWidths() As Double
	Dim regionHeights() As Double
	ReDim regionCenters(1 To outerRegions.Count)
	ReDim regionWidths(1 To outerRegions.Count)
	ReDim regionHeights(1 To outerRegions.Count)
	ReDim regionLabels(1 To outerRegions.Count)
    
	Dim i As Long
	Dim maxWidth As Double: maxWidth = 0
    
	For i = 1 To outerRegions.Count
		Dim reg As AcadRegion
		Set reg = outerRegions(i)
        
		Dim ents As Collection
		Set ents = CollectEntitiesForRegion(reg, shapeSS, allRegions)
		regionEntities.Add ents
		regionLabels(i) = ExtractRegionLabel(ents)
        
		Dim centerPt() As Double
		centerPt = GetEntitySetCenter(reg)
		regionCenters(i) = centerPt
        
		Dim bestHeight As Double
		Dim bestAngle As Double
		bestAngle = FindBestRotationAngleForEntities(ents, centerPt, bestHeight)
		If Abs(bestAngle) > 0.001 Then
			RotateEntities ents, centerPt, bestAngle
		End If
        
		Dim minPt As Variant, maxPt As Variant
		GetEntitiesBounds ents, minPt, maxPt
		regionWidths(i) = maxPt(0) - minPt(0)
		regionHeights(i) = maxPt(1) - minPt(1)
		If regionWidths(i) > maxWidth Then maxWidth = regionWidths(i)
	Next i

	' Sort regions so labeled shapes go first (alphabetical), followed by unlabeled ones
	Set regionEntities = OrderRegionsByLabel(regionEntities, regionLabels)
    
	Dim centers As Collection
	Dim cellWidth As Double, cellHeight As Double
	Dim xGrid() As Double, yGrid() As Double
	Dim cols As Long, rows As Long
	Dim gridSS As AcadSelectionSet
	Set centers = New Collection
	If Not DetectGridFromUserSelection(centers, cellWidth, cellHeight, xGrid, yGrid, cols, rows, gridSS) Then
		GoTo Cleanup
	End If
    
	If maxWidth <= 0 Or cellWidth <= 0 Then
		MsgBox "Falha ao calcular larguras.", vbExclamation, "Distribuir Formas"
		GoTo Cleanup
	End If
    
	' Padding factor to leave extra room inside each cell; adjust to fine-tune fit.
	Dim scalePaddingFactor As Double
	scalePaddingFactor = 1.1
	Dim scaleFactor As Double
	scaleFactor = (maxWidth / cellWidth) * scalePaddingFactor
	
	' Visualize grid centers BEFORE scaling (yellow points for debugging)
	'VisualizeGridCenters doc, centers, xGrid, acYellow, "Before Scaling"
	
	If scaleFactor > 0.000001 Then
		Dim origin(0 To 2) As Double
		Dim gridMin As Variant, gridMax As Variant
		GetSelectionSetBounds gridSS, gridMin, gridMax
		origin(0) = gridMin(0)
		origin(1) = gridMin(1)
		origin(2) = 0
		ScaleEntitiesInSelection gridSS, origin, scaleFactor
		ScaleGridData xGrid, yGrid, origin, scaleFactor
		cellWidth = AverageStep(xGrid)
		cellHeight = AverageStep(yGrid)
		RebuildCentersFromGrid xGrid, yGrid, centers
	End If
    
	' Visualize grid centers AFTER scaling (red points)
	'VisualizeGridCenters doc, centers, xGrid, acRed, "After Scaling"
	DistributeToGrid regionEntities, centers, xGrid, yGrid, cellHeight
    
Cleanup:
	On Error Resume Next
	If Not shapesLayer Is Nothing Then shapesLayer.LayerOn = False
	If Not gridSS Is Nothing Then gridSS.Delete
	shapeSS.Delete
	doc.EndUndoMark
	If Not formPerfisul01 Is Nothing Then formPerfisul01.Show
	Exit Sub
    
ErrHandler:
	MsgBox "Erro: " & Err.Description, vbCritical, "Distribuir Formas"
	Resume Cleanup
End Sub

'-----------------------------
' Shape detection and grouping
'-----------------------------

Private Function DetectOuterRegionsFromSelection(doc As AcadDocument, ss As AcadSelectionSet, ByRef allRegions As Collection) As Collection
	Dim objs() As Object
	ReDim objs(ss.Count - 1)
	Dim i As Long
	For i = 0 To ss.Count - 1
		Set objs(i) = ss.Item(i)
	Next i
    
	Dim created As Variant
	On Error Resume Next
	created = doc.ModelSpace.AddRegion(objs)
	If Err.Number <> 0 Or IsEmpty(created) Then
		Err.Clear
		Exit Function
	End If
	On Error GoTo 0
    
	Dim regArr() As AcadRegion
	Dim idx As Long
	ReDim regArr(LBound(created) To UBound(created))
	idx = LBound(created)
	Dim r As Variant
	For Each r In created
		Set regArr(idx) = r
		idx = idx + 1
	Next r

	Dim shapesLayer As AcadLayer
	Set shapesLayer = EnsureShapesLayer(doc)
	For i = LBound(regArr) To UBound(regArr)
		On Error Resume Next
		regArr(i).Layer = "Shapes"
		regArr(i).Color = acByLayer
		Err.Clear
		On Error GoTo 0
	Next i

	Set allRegions = New Collection
	For i = LBound(regArr) To UBound(regArr)
		allRegions.Add regArr(i)
	Next i

	Set DetectOuterRegionsFromSelection = FilterOuterRegions(regArr)
End Function

Private Function FilterOuterRegions(regArr() As AcadRegion) As Collection
	Dim count As Long
	count = UBound(regArr) - LBound(regArr) + 1
	If count = 0 Then Exit Function
    
	Dim i As Long, j As Long
	Dim swapped As Boolean
	Dim tmp As AcadRegion
	' Sort by area desc
	Do
		swapped = False
		For i = LBound(regArr) To UBound(regArr) - 1
			If regArr(i).Area < regArr(i + 1).Area Then
				Set tmp = regArr(i)
				Set regArr(i) = regArr(i + 1)
				Set regArr(i + 1) = tmp
				swapped = True
			End If
		Next i
	Loop While swapped
    
	Dim keepFlags() As Boolean
	ReDim keepFlags(LBound(regArr) To UBound(regArr))
	For i = LBound(regArr) To UBound(regArr)
		keepFlags(i) = True
	Next i
    
	Dim testReg As AcadRegion, containerReg As AcadRegion
	Dim copyA As AcadRegion, copyB As AcadRegion
	For i = LBound(regArr) To UBound(regArr)
		If keepFlags(i) Then
			Set testReg = regArr(i)
			For j = LBound(regArr) To UBound(regArr)
				If i <> j And keepFlags(j) Then
					Set containerReg = regArr(j)
					If containerReg.Area >= testReg.Area Then
						On Error Resume Next
						Set copyA = testReg.Copy
						Set copyB = containerReg.Copy
						copyA.Boolean acIntersection, copyB
						If Err.Number = 0 Then
							If Abs(copyA.Area - testReg.Area) < 0.0001 Then
								keepFlags(i) = False
								copyA.Delete
								copyB.Delete
								Exit For
							End If
						End If
						If Not copyA Is Nothing Then copyA.Delete
						If Not copyB Is Nothing Then copyB.Delete
						Err.Clear
						On Error GoTo 0
					End If
				End If
			Next j
		End If
	Next i
    
	Dim result As New Collection
	For i = LBound(regArr) To UBound(regArr)
		If keepFlags(i) Then result.Add regArr(i)
	Next i
	Set FilterOuterRegions = result
End Function

Private Function CollectEntitiesForRegion(reg As AcadRegion, ss As AcadSelectionSet, allRegions As Collection) As Collection
	Dim col As New Collection
	Dim regMin As Variant, regMax As Variant
	reg.GetBoundingBox regMin, regMax
    
	Dim regCx As Double, regCy As Double
	regCx = (regMin(0) + regMax(0)) / 2
	regCy = (regMin(1) + regMax(1)) / 2
    
	Dim ent As AcadEntity
	For Each ent In ss
		Dim eMin As Variant, eMax As Variant
		On Error Resume Next
		ent.GetBoundingBox eMin, eMax
		If Err.Number = 0 Then
			Dim cx As Double, cy As Double
			cx = (eMin(0) + eMax(0)) / 2
			cy = (eMin(1) + eMax(1)) / 2
			If cx >= regMin(0) - 0.01 And cx <= regMax(0) + 0.01 And _
			   cy >= regMin(1) - 0.01 And cy <= regMax(1) + 0.01 Then
				col.Add ent
			End If
		End If
		Err.Clear
		On Error GoTo 0
	Next ent

	If Not allRegions Is Nothing Then
		Dim regEnt As AcadRegion
		For Each regEnt In allRegions
			If regEnt Is reg Then GoTo NextReg
			Dim rMin As Variant, rMax As Variant
			On Error Resume Next
			regEnt.GetBoundingBox rMin, rMax
			If Err.Number = 0 Then
				Dim rcx As Double, rcy As Double
				rcx = (rMin(0) + rMax(0)) / 2
				rcy = (rMin(1) + rMax(1)) / 2
				If rcx >= regMin(0) - 0.01 And rcx <= regMax(0) + 0.01 And _
				   rcy >= regMin(1) - 0.01 And rcy <= regMax(1) + 0.01 Then
					col.Add regEnt
				End If
			End If
			Err.Clear
			On Error GoTo 0
NextReg:
		Next regEnt
	End If
	col.Add reg
	Set CollectEntitiesForRegion = col
End Function

Private Function GetEntitySetCenter(entObj As AcadEntity) As Double()
	Dim minPt As Variant, maxPt As Variant
	entObj.GetBoundingBox minPt, maxPt
	Dim c(0 To 2) As Double
	c(0) = (minPt(0) + maxPt(0)) / 2
	c(1) = (minPt(1) + maxPt(1)) / 2
	c(2) = 0
	GetEntitySetCenter = c
End Function

'-----------------------------
' Orientation helpers
'-----------------------------

Private Function FindBestRotationAngleForEntities(ents As Collection, centerPt() As Double, ByRef heightOut As Double) As Double
	Const STEP_DEG As Double = 1
	Const PI As Double = 3.14159265358979
    
	Dim points() As Point2D
	Dim numPoints As Long
	numPoints = CollectSamplingPointsFromCollection(ents, centerPt, points)
    
	Dim bestAngle As Double: bestAngle = 0
	Dim bestHeight As Double: bestHeight = 1E+30
	Dim bestAspect As Double: bestAspect = 0
    
	Dim deg As Double, angle As Double
	Dim width As Double, height As Double, aspect As Double
	For deg = 0 To 180 Step STEP_DEG
		angle = deg * PI / 180
		GetRotatedBoundsFromPoints points, numPoints, angle, width, height
		If height > 0 Then aspect = width / height Else aspect = 0
		If height < bestHeight - 0.001 Or (Abs(height - bestHeight) < 0.001 And aspect > bestAspect) Then
			bestHeight = height
			bestAspect = aspect
			bestAngle = angle
		End If
	Next deg
	heightOut = bestHeight
	FindBestRotationAngleForEntities = bestAngle
End Function

Private Function CollectSamplingPointsFromCollection(ents As Collection, centerPt() As Double, ByRef pointsOut() As Point2D) As Long
	Dim ent As AcadEntity
	Dim count As Long: count = 0
	ReDim pointsOut(0 To 1000) As Point2D
	For Each ent In ents
		If IsExcludedFromAngleCalculation(ent) Then GoTo NextEnt
		Dim name As String
		name = UCase$(ent.ObjectName)
		If InStr(1, name, "POLYLINE", vbTextCompare) > 0 Then
			CollectPolylinePoints ent, centerPt, pointsOut, count
		ElseIf InStr(1, name, "REGION", vbTextCompare) > 0 Then
			CollectRegionPoints ent, centerPt, pointsOut, count
		ElseIf name = "ACDBLINE" Or name = "ACADLINE" Then
			CollectLinePoints ent, centerPt, pointsOut, count
		ElseIf name = "ACDBARC" Or name = "ACADARC" Then
			CollectArcPoints ent, centerPt, pointsOut, count
		Else
			CollectBoundingBoxPoints ent, centerPt, pointsOut, count
		End If
NextEnt:
	Next ent
	CollectSamplingPointsFromCollection = count
End Function

Private Function IsExcludedFromAngleCalculation(ent As AcadEntity) As Boolean
	Dim nm As String
	nm = UCase$(ent.ObjectName)
	If InStr(1, nm, "REGION", vbTextCompare) > 0 Then
		IsExcludedFromAngleCalculation = True
		Exit Function
	End If
	Dim layerName As String
	layerName = UCase$(Trim$(ent.Layer))
	If layerName = "SHAPES" Then
		IsExcludedFromAngleCalculation = True
		Exit Function
	End If
	IsExcludedFromAngleCalculation = False
End Function

Private Sub RotateEntities(ents As Collection, centerPt() As Double, angle As Double)
	Dim ent As AcadEntity
	For Each ent In ents
		On Error Resume Next
		ent.Rotate centerPt, angle
		Err.Clear
		On Error GoTo 0
	Next ent
End Sub

Private Sub ScaleEntities(ents As Collection, centerPt() As Double, scaleFactor As Double)
	If Abs(scaleFactor - 1) < 0.0001 Then Exit Sub
	Dim ent As AcadEntity
	For Each ent In ents
		On Error Resume Next
		ent.ScaleEntity centerPt, scaleFactor
		Err.Clear
		On Error GoTo 0
	Next ent
End Sub

Private Sub ScaleEntitiesInSelection(ss As AcadSelectionSet, origin() As Double, scaleFactor As Double)
	If ss Is Nothing Then Exit Sub
	If Abs(scaleFactor - 1) < 0.0001 Then Exit Sub
	Dim ent As AcadEntity
	For Each ent In ss
		On Error Resume Next
		ent.ScaleEntity origin, scaleFactor
		Err.Clear
		On Error GoTo 0
	Next ent
End Sub

Private Sub GetEntitiesBounds(ents As Collection, ByRef minPt As Variant, ByRef maxPt As Variant)
	' Calculate the bounding box encompassing all entities in the collection
	Dim minX As Double, minY As Double, maxX As Double, maxY As Double
	Dim first As Boolean: first = True
	
	Dim ent As AcadEntity
	Dim eMin As Variant, eMax As Variant
	
	' Iterate through each entity to find overall min/max coordinates
	For Each ent In ents
		On Error Resume Next
		' Get individual entity's bounding box
		ent.GetBoundingBox eMin, eMax
		If Err.Number = 0 Then
			' On first entity, initialize bounds
			If first Then
				minX = eMin(0): minY = eMin(1): maxX = eMax(0): maxY = eMax(1)
				first = False
			Else
				' On subsequent entities, expand bounds if needed
				If eMin(0) < minX Then minX = eMin(0)
				If eMin(1) < minY Then minY = eMin(1)
				If eMax(0) > maxX Then maxX = eMax(0)
				If eMax(1) > maxY Then maxY = eMax(1)
			End If
		End If
		Err.Clear
		On Error GoTo 0
	Next ent
	
	' Return bounding box as [minX, minY, Z] and [maxX, maxY, Z]
	minPt = Array(minX, minY, 0)
	maxPt = Array(maxX, maxY, 0)
End Sub

Private Sub GetSelectionSetBounds(ss As AcadSelectionSet, ByRef minPt As Variant, ByRef maxPt As Variant)
	' Bounding box for a selection set
	Dim minX As Double, minY As Double, maxX As Double, maxY As Double
	Dim first As Boolean: first = True
	Dim ent As AcadEntity
	Dim eMin As Variant, eMax As Variant
	For Each ent In ss
		On Error Resume Next
		ent.GetBoundingBox eMin, eMax
		If Err.Number = 0 Then
			If first Then
				minX = eMin(0): minY = eMin(1): maxX = eMax(0): maxY = eMax(1)
				first = False
			Else
				If eMin(0) < minX Then minX = eMin(0)
				If eMin(1) < minY Then minY = eMin(1)
				If eMax(0) > maxX Then maxX = eMax(0)
				If eMax(1) > maxY Then maxY = eMax(1)
			End If
		End If
		Err.Clear
		On Error GoTo 0
	Next ent
	minPt = Array(minX, minY, 0)
	maxPt = Array(maxX, maxY, 0)
End Sub

Private Sub MoveEntities(ents As Collection, fromPt() As Double, toPt() As Double)
	' Move all entities from one point to another
	Dim ent As AcadEntity
	For Each ent In ents
		On Error Resume Next
		' Move the entity: calculates translation vector and applies to all geometry
		ent.Move fromPt, toPt
		Err.Clear
		On Error GoTo 0
	Next ent
End Sub

Private Sub ScaleGridData(ByRef xGrid() As Double, ByRef yGrid() As Double, origin() As Double, scaleFactor As Double)
	Dim i As Long
	For i = 0 To UBound(xGrid)
		xGrid(i) = origin(0) + (xGrid(i) - origin(0)) * scaleFactor
	Next i
	For i = 0 To UBound(yGrid)
		yGrid(i) = origin(1) + (yGrid(i) - origin(1)) * scaleFactor
	Next i
End Sub

'-----------------------------
' Grid detection
'-----------------------------

Private Function DetectGridFromUserSelection(centers As Collection, ByRef cellWidth As Double, ByRef cellHeight As Double, _
	ByRef xGrid() As Double, ByRef yGrid() As Double, ByRef cols As Long, ByRef rows As Long, _
	ByRef gridSS As AcadSelectionSet) As Boolean
	
	' Radius offset multipliers for grid corner calculation
	Const HORIZONTAL_RADIUS_MULTIPLIER As Double = 1.2
	Const VERTICAL_RADIUS_MULTIPLIER As Double = 1.2

	Dim hasTopRightCorner As Boolean
	Dim topRightX As Double, topRightY As Double
	
	Dim sset As AcadSelectionSet
	Set sset = PrepareSelectionSet(ThisDrawing, "DSG_GRID")
	MsgBox "Selecione a área cobrindo as linhas da grade.", vbInformation, "Distribuir Formas"
	sset.SelectOnScreen
	If sset.Count = 0 Then Exit Function
    
	Dim xArr() As Double, yArr() As Double
	ReDim xArr(0 To 200)
	ReDim yArr(0 To 200)
	Dim xCount As Long, yCount As Long
	xCount = 0: yCount = 0
	
	' Find circles to determine grid extent
	Dim circles As Collection
	Set circles = New Collection
	Dim circleRadii As Collection
	Set circleRadii = New Collection
    
	Dim ent As AcadEntity
	For Each ent In sset
		If ent.ObjectName = "AcDbLine" Then
			Dim ln As AcadLine
			Set ln = ent
			Dim sx As Double, sy As Double, ex As Double, ey As Double
			sx = ln.StartPoint(0): sy = ln.StartPoint(1)
			ex = ln.EndPoint(0): ey = ln.EndPoint(1)
			If Abs(sx - ex) < 0.1 Then AddUniqueVal xArr, xCount, (sx + ex) / 2
			If Abs(sy - ey) < 0.1 Then AddUniqueVal yArr, yCount, (sy + ey) / 2
		ElseIf ent.ObjectName = "AcDbCircle" Then
			Dim circ As AcadCircle
			Set circ = ent
			Dim circCenter As Variant
			circCenter = circ.center
			circles.Add circCenter
			circleRadii.Add circ.radius
		End If
	Next ent
    
	If xCount < 2 Or yCount < 2 Then
		MsgBox "Grade inválida. Verticais: " & xCount & " horizontais: " & yCount, vbExclamation, "Distribuir Formas"
		sset.Delete
		Exit Function
	End If
    
	SortDoubles xArr, xCount
	SortDoubles yArr, yCount
	xCount = MergeCloseSorted(xArr, xCount)
	yCount = MergeCloseSorted(yArr, yCount)
	
	' If circles were found, use them to determine grid extent
	If circles.Count > 0 Then
		' Sort circles by Y (descending, top to bottom) then by X (ascending, left to right)
		Dim i As Long, j As Long
		Dim sortedCircles As Collection
		Dim sortedRadii As Collection
		Set sortedCircles = New Collection
		Set sortedRadii = New Collection
		
		' Simple bubble sort by Y descending, then X ascending
		Dim circArray() As Variant
		ReDim circArray(1 To circles.Count)
		Dim radArray() As Double
		ReDim radArray(1 To circles.Count)
		
		For i = 1 To circles.Count
			circArray(i) = circles(i)
			radArray(i) = circleRadii(i)
		Next i
		
		Dim swapped As Boolean
		For i = 1 To circles.Count - 1
			swapped = False
			For j = 1 To circles.Count - i
				Dim y1 As Double, y2 As Double, x1 As Double, x2 As Double
				y1 = CDbl(circArray(j)(1))
				y2 = CDbl(circArray(j + 1)(1))
				x1 = CDbl(circArray(j)(0))
				x2 = CDbl(circArray(j + 1)(0))
				
				If y1 < y2 Or (Abs(y1 - y2) < 0.1 And x1 > x2) Then
					Dim tmpArr As Variant
					Dim tmpRad As Double
					tmpArr = circArray(j)
					tmpRad = radArray(j)
					circArray(j) = circArray(j + 1)
					radArray(j) = radArray(j + 1)
					circArray(j + 1) = tmpArr
					radArray(j + 1) = tmpRad
					swapped = True
				End If
			Next j
			If Not swapped Then Exit For
		Next i
		
		' Find the rightmost circle in the first row
		Dim firstRowY As Double
		firstRowY = CDbl(circArray(1)(1))
		Dim lastCircleInFirstRow As Variant
		Dim lastRadiusInFirstRow As Double
		Dim circleCount As Long: circleCount = 0
		
		For i = 1 To circles.Count
			If Abs(CDbl(circArray(i)(1)) - firstRowY) < 0.1 Then
				lastCircleInFirstRow = circArray(i)
				lastRadiusInFirstRow = radArray(i)
				circleCount = i
			Else
				Exit For
			End If
		Next i
		
		' Calculate grid extent from the last circle in first row
		topRightX = CDbl(lastCircleInFirstRow(0)) + lastRadiusInFirstRow * HORIZONTAL_RADIUS_MULTIPLIER
		topRightY = CDbl(lastCircleInFirstRow(1)) + lastRadiusInFirstRow * VERTICAL_RADIUS_MULTIPLIER
		hasTopRightCorner = True

		' Derive vertical cell height from first circle in first and second rows
		Dim circleHeightStep As Double
		If circles.Count >= circleCount + 1 And circleCount > 0 Then
			circleHeightStep = Abs(CDbl(circArray(1)(1)) - CDbl(circArray(circleCount + 1)(1)))
		End If
		
		' Adjust grid arrays to match circle positions
		' xArr should represent the column lines
		' yArr should represent the row lines
		' Use the circles themselves to refine the grid
		cols = circleCount - 1
		rows = circles.Count \ circleCount - 1
	End If
	
	ReDim xGrid(0 To xCount - 1)
	ReDim yGrid(0 To yCount - 1)
	Dim k As Long
	For k = 0 To xCount - 1: xGrid(k) = xArr(k): Next k
	For k = 0 To yCount - 1: yGrid(k) = yArr(k): Next k
    
	If cols = 0 Then cols = xCount - 1
	If rows = 0 Then rows = yCount - 1
	cellWidth = AverageStep(xGrid)
	cellHeight = AverageStep(yGrid)

	' If we obtained a circle-based vertical step, prefer it for consistent row spacing
	If circleHeightStep > 0.000001 Then cellHeight = circleHeightStep

	' Anchor top-right corner from the last circle in first row (circle-based extents)
	If hasTopRightCorner Then
		If yCount > 0 Then yGrid(yCount - 1) = topRightY
		If xCount > 0 Then xGrid(xCount - 1) = topRightX
	End If

	' Normalize all Y grid lines to be equally spaced so every row step equals the averaged cellHeight
	If yCount >= 2 Then
		Dim topY As Double
		topY = yGrid(yCount - 1)
		For i = 0 To yCount - 1
			yGrid(yCount - 1 - i) = topY - cellHeight * i
		Next i
	End If
    
	BuildCenters xGrid, yGrid, centers
	Set gridSS = sset
	DetectGridFromUserSelection = True
End Function

Private Sub BuildCenters(xGrid() As Double, yGrid() As Double, centers As Collection)
	' Generate center points for all grid cells by iterating rows and columns
	Dim r As Long, c As Long
	
	' Iterate rows from top to bottom (reverse order: highest Y first)
	For r = UBound(yGrid) - 1 To 0 Step -1
		' Iterate columns from left to right
		For c = 0 To UBound(xGrid) - 1
			' Create center point for this cell
			Dim pt(0 To 2) As Double
			
			' Horizontal center: midpoint between adjacent X grid lines
			pt(0) = (xGrid(c) + xGrid(c + 1)) / 2
			
			' Vertical center: midpoint between adjacent Y grid lines
			pt(1) = (yGrid(r) + yGrid(r + 1)) / 2
			
			' Z coordinate (depth)
			pt(2) = 0
			
			' Add this cell center to collection
			centers.Add pt
		Next c
	Next r
End Sub

Private Sub RebuildCentersFromGrid(xGrid() As Double, yGrid() As Double, centers As Collection)
	' Clear existing centers and rebuild from current grid state
	' This is used after scaling the grid to recalculate all cell centers
	Do While centers.Count > 0
		centers.Remove 1
	Loop
	BuildCenters xGrid, yGrid, centers
End Sub

Private Function AverageStep(arr() As Double) As Double
	Dim i As Long, total As Double
	For i = 0 To UBound(arr) - 1
		total = total + (arr(i + 1) - arr(i))
	Next i
	If UBound(arr) > 0 Then
		AverageStep = total / (UBound(arr))
	Else
		AverageStep = 0
	End If
End Function

Private Sub AddUniqueVal(ByRef arr() As Double, ByRef count As Long, val As Double)
	Dim i As Long
	For i = 0 To count - 1
		If Abs(arr(i) - val) < 0.1 Then Exit Sub
	Next i
	If count > UBound(arr) Then ReDim Preserve arr(0 To UBound(arr) + 50)
	arr(count) = val
	count = count + 1
End Sub

Private Sub SortDoubles(ByRef arr() As Double, count As Long)
	Dim i As Long, j As Long, tmp As Double
	For i = 0 To count - 2
		For j = i + 1 To count - 1
			If arr(i) > arr(j) Then
				tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
			End If
		Next j
	Next i
End Sub

Private Function MergeCloseSorted(ByRef arr() As Double, count As Long) As Long
	If count <= 1 Then
		MergeCloseSorted = count
		Exit Function
	End If
	Dim minStep As Double: minStep = 0
	Dim i As Long
	For i = 0 To count - 2
		Dim d As Double
		d = Abs(arr(i + 1) - arr(i))
		If d > 0.000001 Then
			If minStep = 0 Or d < minStep Then minStep = d
		End If
	Next i
	If minStep = 0 Then
		MergeCloseSorted = 1
		arr(0) = arr(0)
		Exit Function
	End If
	Dim tol As Double
	tol = minStep * 1.05
	Dim writeIdx As Long: writeIdx = 0
	For i = 0 To count - 1
		If writeIdx = 0 Then
			arr(writeIdx) = arr(i)
			writeIdx = writeIdx + 1
		Else
			If Abs(arr(i) - arr(writeIdx - 1)) > tol Then
				arr(writeIdx) = arr(i)
				writeIdx = writeIdx + 1
			End If
		End If
	Next i
	MergeCloseSorted = writeIdx
End Function

'-----------------------------
' Distribution
'-----------------------------

Private Sub DistributeToGrid(regionEntities As Collection, centers As Collection, xGrid() As Double, yGrid() As Double, cellHeight As Double)
	' Vertical positioning fine-tuning: adjust this offset to move all extra rows up/down
	' Positive value = move down, Negative value = move up
	Const VERTICAL_OFFSET As Double = 0
	
	' Calculate if additional rows are needed beyond the initial grid
	Dim totalCells As Long
	totalCells = centers.Count
	Dim need As Long
	need = regionEntities.Count - totalCells
	Dim extraRows As Long
	
	' If there are more shapes than grid cells, create additional rows below the grid
	If need > 0 Then
		Dim cols As Long
		cols = UBound(xGrid)
		extraRows = (need + cols - 1) \ cols
		AppendExtraRows centers, xGrid, yGrid(0), cellHeight, extraRows, VERTICAL_OFFSET
	End If
    
	' Move each shape to its assigned grid cell center
	Dim i As Long
	For i = 1 To regionEntities.Count
		' Check if we have a corresponding center point for this shape
		If i > centers.Count Then Exit For
		
		' Get target center point from grid
		Dim tgt() As Double
		tgt = centers(i)
		
		' Calculate current geometric center of the shape from its bounding box
		Dim minPt As Variant, maxPt As Variant
		GetEntitiesBounds regionEntities(i), minPt, maxPt
		
		' Extract current center X and Y coordinates
		Dim curr(0 To 2) As Double
		curr(0) = (minPt(0) + maxPt(0)) / 2
		curr(1) = (minPt(1) + maxPt(1)) / 2
		curr(2) = 0
		
		' Move shape from current center to target grid cell center
		MoveEntities regionEntities(i), curr, tgt
	Next i
End Sub

Private Sub AppendExtraRows(centers As Collection, xGrid() As Double, baseY As Double, cellHeight As Double, extraRows As Long, Optional verticalOffset As Double = 0)
	' Generate center points for additional grid rows needed beyond the initial grid
	Dim r As Long, c As Long
	
	' For each extra row to be created
	For r = 1 To extraRows
		' For each column in the grid
		For c = 0 To UBound(xGrid) - 1
			' Create a new center point for this row/column intersection
			Dim pt(0 To 2) As Double
			
			' Horizontal: center of the cell width
			pt(0) = (xGrid(c) + xGrid(c + 1)) / 2
			
			' Vertical: position below the base grid by (r * cellHeight) with offset to cell center
			' Fine-tuned formula: baseY - (cellHeight * r) + (cellHeight / 2) + verticalOffset
			' verticalOffset allows adjustment of all extra rows relative to base grid
			pt(1) = baseY - (cellHeight * r) + (cellHeight / 2) + verticalOffset
			
			' Z coordinate (depth)
			pt(2) = 0
			
			' Add this center point to the collection
			centers.Add pt
		Next c
	Next r
End Sub

Private Sub VisualizeGridCenters(doc As AcadDocument, centers As Collection, xGrid() As Double, Optional pointColor As Long = acRed, Optional debugLabel As String = "")
	' Draw points at grid cell centers with hyperlink metadata containing cell coordinates
	' pointColor: color for the points (default acRed)
	' debugLabel: optional label prefix for hyperlink descriptions (e.g., "Before Scaling", "After Scaling")
	On Error Resume Next
	Dim i As Long
	Dim pt As Variant
	Dim cols As Long
	Dim cellIndex As Long
	Dim row As Long, col As Long
	Dim label As String
	Dim hyperLabel As String
	Dim pointObj As Object
	
	cols = UBound(xGrid)
	
	For i = 1 To centers.Count
		pt = centers(i)
		
		' Calculate row and column from index
		cellIndex = i - 1
		row = cellIndex \ cols
		col = cellIndex Mod cols
		
		' Format label as R##C##
		label = "R" & Format(row + 1, "00") & "C" & Format(col + 1, "00")
		
		' Include debug label in hyperlink description if provided
		If debugLabel <> "" Then
			hyperLabel = debugLabel & " - " & label
		Else
			hyperLabel = label
		End If
		
		' Create point at grid center
		Set pointObj = doc.ModelSpace.AddPoint(pt)
		
		' Set color and add hyperlink metadata
		If Not pointObj Is Nothing Then
			pointObj.Color = pointColor
			
			' Add hyperlink metadata with cell coordinate information
			If pointObj.Hyperlinks.Count >= 0 Then
				pointObj.Hyperlinks.Add hyperLabel
				pointObj.Hyperlinks(1).Description = hyperLabel
			End If
		End If
		
		Err.Clear
	Next i
	
	doc.Application.Refresh
	Err.Clear
	On Error GoTo 0
End Sub

'-----------------------------
' Geometry collection helpers (borrowed from autoOrient)
'-----------------------------

Private Sub AddPoint(x As Double, y As Double, ByRef points() As Point2D, ByRef count As Long)
	If count > UBound(points) Then ReDim Preserve points(0 To UBound(points) * 2) As Point2D
	points(count).x = x
	points(count).y = y
	count = count + 1
End Sub

Private Sub CollectRegionPoints(ent As AcadEntity, centerPt() As Double, ByRef points() As Point2D, ByRef count As Long)
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
		Dim nm As String
		nm = UCase$(subEnt.ObjectName)
		If InStr(1, nm, "REGION", vbTextCompare) > 0 Then
			CollectBoundingBoxPoints subEnt, centerPt, points, count
		ElseIf InStr(1, nm, "LINE", vbTextCompare) > 0 Then
			CollectLinePoints subEnt, centerPt, points, count
		ElseIf InStr(1, nm, "ARC", vbTextCompare) > 0 Then
			CollectArcPoints subEnt, centerPt, points, count
		Else
			CollectBoundingBoxPoints subEnt, centerPt, points, count
		End If
		subEnt.Delete
	Next i
	On Error GoTo 0
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
		Dim nm As String
		nm = UCase$(subEnt.ObjectName)
		If InStr(1, nm, "LINE", vbTextCompare) > 0 Then
			CollectLinePoints subEnt, centerPt, points, count
		ElseIf InStr(1, nm, "ARC", vbTextCompare) > 0 Then
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
	Dim sp As Variant, ep As Variant
	sp = lineEnt.StartPoint
	ep = lineEnt.EndPoint
	AddPoint sp(0) - centerPt(0), sp(1) - centerPt(1), points, count
	AddPoint ep(0) - centerPt(0), ep(1) - centerPt(1), points, count
	On Error GoTo 0
End Sub

Private Sub CollectArcPoints(arcEnt As AcadEntity, centerPt() As Double, ByRef points() As Point2D, ByRef count As Long)
	On Error Resume Next
	Dim sp As Variant, ep As Variant
	sp = arcEnt.StartPoint
	ep = arcEnt.EndPoint
	AddPoint sp(0) - centerPt(0), sp(1) - centerPt(1), points, count
	AddPoint ep(0) - centerPt(0), ep(1) - centerPt(1), points, count
	Dim radius As Double, cen As Variant, sa As Double, ea As Double
	radius = arcEnt.radius
	cen = arcEnt.center
	sa = arcEnt.startAngle
	ea = arcEnt.endAngle
	Dim diff As Double
	diff = ea - sa
	If diff <= 0 Then diff = diff + 6.28318530717959
	Dim mid As Double
	mid = sa + diff / 2
	Dim mx As Double, my As Double
	mx = cen(0) + radius * Cos(mid)
	my = cen(1) + radius * Sin(mid)
	AddPoint mx - centerPt(0), my - centerPt(1), points, count
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

Private Sub GetRotatedBoundsFromPoints(points() As Point2D, count As Long, angle As Double, ByRef widthOut As Double, ByRef heightOut As Double)
	Const LARGE As Double = 1E+30
	Dim minX As Double: minX = LARGE
	Dim minY As Double: minY = LARGE
	Dim maxX As Double: maxX = -LARGE
	Dim maxY As Double: maxY = -LARGE
	Dim cosA As Double: cosA = Cos(angle)
	Dim sinA As Double: sinA = Sin(angle)
	Dim i As Long, rx As Double, ry As Double
	For i = 0 To count - 1
		rx = points(i).x * cosA - points(i).y * sinA
		ry = points(i).x * sinA + points(i).y * cosA
		If rx < minX Then minX = rx
		If rx > maxX Then maxX = rx
		If ry < minY Then minY = ry
		If ry > maxY Then maxY = ry
	Next i
	If minX > maxX Then
		widthOut = 0: heightOut = 0
	Else
		widthOut = maxX - minX
		heightOut = maxY - minY
	End If
End Sub

'-----------------------------
' Utilities
'-----------------------------

Private Function ExtractRegionLabel(ents As Collection) As String
	Dim ent As AcadEntity
	For Each ent In ents
		If TypeOf ent Is AcadText Or TypeOf ent Is AcadMText Then
			If IsGravacaoLayerName(ent.Layer) Then
				Dim txt As String
				txt = Trim$(GetEntityTextString(ent))
				If txt <> "" Then
					If Not ContainsExcludedKeyword(txt) Then
						ExtractRegionLabel = txt
						Exit Function
					End If
				End If
			End If
		End If
	Next ent
End Function

Private Function GetEntityTextString(ent As AcadEntity) As String
	On Error Resume Next
	If TypeOf ent Is AcadText Then
		GetEntityTextString = ent.TextString
	ElseIf TypeOf ent Is AcadMText Then
		GetEntityTextString = ent.TextString
	End If
	Err.Clear
	On Error GoTo 0
End Function

Private Function IsGravacaoLayerName(layerName As String) As Boolean
	Dim normalized As String
	normalized = LCase$(StripDiacritics(Trim$(layerName)))
	IsGravacaoLayerName = (normalized = "gravacao")
End Function

Private Function ContainsExcludedKeyword(textVal As String) As Boolean
	Dim normalized As String
	normalized = LCase$(StripDiacritics(textVal))
	If InStr(1, normalized, "carroceria", vbTextCompare) > 0 Then ContainsExcludedKeyword = True: Exit Function
	If InStr(1, normalized, "ferramentaria", vbTextCompare) > 0 Then ContainsExcludedKeyword = True: Exit Function
	If InStr(1, normalized, "portas", vbTextCompare) > 0 Then ContainsExcludedKeyword = True: Exit Function
	If InStr(1, normalized, "tampa", vbTextCompare) > 0 Then ContainsExcludedKeyword = True: Exit Function
	If InStr(1, normalized, "vidros", vbTextCompare) > 0 Then ContainsExcludedKeyword = True: Exit Function
	If InStr(1, normalized, "teto", vbTextCompare) > 0 Then ContainsExcludedKeyword = True
End Function

Private Function OrderRegionsByLabel(regionEntities As Collection, labels() As String) As Collection
	Dim labeled As New Collection
	Dim unlabeled As New Collection
	Dim i As Long
	For i = 1 To regionEntities.Count
		If Trim$(labels(i)) <> "" Then
			labeled.Add i
		Else
			unlabeled.Add i
		End If
	Next i

	Dim sortedLabeled As Collection
	Set sortedLabeled = SortIndicesByLabel(labels, labeled)

	Dim result As New Collection
	Dim idxVar As Variant
	For Each idxVar In sortedLabeled
		result.Add regionEntities(idxVar)
	Next idxVar
	For Each idxVar In unlabeled
		result.Add regionEntities(idxVar)
	Next idxVar
	Set OrderRegionsByLabel = result
End Function

Private Function SortIndicesByLabel(labels() As String, indices As Collection) As Collection
	Dim count As Long
	count = indices.Count
	If count = 0 Then
		Dim emptyCol As New Collection
		Set SortIndicesByLabel = emptyCol
		Exit Function
	End If
	Dim arr() As Long
	ReDim arr(1 To count)
	Dim i As Long, j As Long
	For i = 1 To count
		arr(i) = CLng(indices(i))
	Next i
	For i = 1 To count - 1
		For j = i + 1 To count
			If StrComp(labels(arr(i)), labels(arr(j)), vbTextCompare) > 0 Then
				Dim tmp As Long
				tmp = arr(i)
				arr(i) = arr(j)
				arr(j) = tmp
			End If
		Next j
	Next i
	Dim sorted As New Collection
	For i = 1 To count
		sorted.Add arr(i)
	Next i
	Set SortIndicesByLabel = sorted
End Function

Private Function PrepareSelectionSet(doc As AcadDocument, name As String) As AcadSelectionSet
	On Error Resume Next
	Dim ss As AcadSelectionSet
	Set ss = doc.SelectionSets.Item(name)
	If Err.Number = 0 Then
		ss.Clear
	Else
		Err.Clear
		Set ss = doc.SelectionSets.Add(name)
	End If
	On Error GoTo 0
	Set PrepareSelectionSet = ss
End Function

Private Sub NormalizeSelectedLayers(ss As AcadSelectionSet)
	Dim ent As AcadEntity
	For Each ent In ss
		Dim layerName As String
		layerName = LCase$(StripDiacritics(Trim$(ent.Layer)))
		If InStr(1, layerName, "gravacao", vbTextCompare) = 0 And _
		   InStr(1, layerName, "dobra", vbTextCompare) = 0 Then
			On Error Resume Next
			ent.Layer = "0"
			Err.Clear
			On Error GoTo 0
		End If
	Next ent
End Sub

Private Function StripDiacritics(ByVal text As String) As String
	Dim i As Long
	Dim ch As String
	Dim code As Long
	Dim result As String
	result = ""
	For i = 1 To Len(text)
		ch = Mid$(text, i, 1)
		code = AscW(ch)
		Select Case code
			Case 192, 193, 194, 195, 196, 197, 224, 225, 226, 227, 228, 229
				result = result & "a"
			Case 199, 231
				result = result & "c"
			Case 200, 201, 202, 203, 232, 233, 234, 235
				result = result & "e"
			Case 204, 205, 206, 207, 236, 237, 238, 239
				result = result & "i"
			Case 210, 211, 212, 213, 214, 242, 243, 244, 245, 246
				result = result & "o"
			Case 217, 218, 219, 220, 249, 250, 251, 252
				result = result & "u"
			Case Else
				result = result & ch
		End Select
	Next i
	StripDiacritics = result
End Function

Private Function EnsureShapesLayer(doc As AcadDocument) As AcadLayer
	Dim shapesLayer As AcadLayer
	On Error Resume Next
	Set shapesLayer = doc.Layers.Item("Shapes")
	If Err.Number <> 0 Then
		Err.Clear
		Set shapesLayer = doc.Layers.Add("Shapes")
		shapesLayer.Color = acGreen
	End If
	On Error GoTo 0
	Set EnsureShapesLayer = shapesLayer
End Function
