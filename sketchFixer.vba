Public Sub FixSmallGaps()
	' Prompts user for min and max gap values, finds endpoints of lines/arcs/curves within that range,
	' zooms in, asks user to fix, and moves endpoints to midpoint if confirmed.
	Dim minGap As Double, maxGap As Double
	Dim doc As AcadDocument
	Dim ent As AcadEntity, ents As Variant
	Dim i As Long, j As Long
	Dim endpoints() As Variant, types() As String
	Dim idx As Long, idx2 As Long
	Dim pt1 As Variant, pt2 As Variant
	Dim dist As Double, midPt(0 To 2) As Double
	Dim resp As VbMsgBoxResult
    
	Set doc = ThisDrawing
    
	' Ask user about unit handling preference
	Dim unitChoice As VbMsgBoxResult
	unitChoice = MsgBox("How should drawing units be handled?" & vbCrLf & vbCrLf & _
	                   "YES = Assume drawing is in METERS" & vbCrLf & _
	                   "NO = Try to detect units automatically" & vbCrLf & _
	                   "CANCEL = Exit", vbYesNoCancel + vbQuestion, "Unit Selection")
	
	If unitChoice = vbCancel Then Exit Sub
    
	' Get drawing units to properly convert input values
	Dim unitFactor As Double
	Dim unitName As String
	
	If unitChoice = vbYes Then
		' User chose to assume meters
		unitFactor = 0.001 ' Convert mm input to meters
		unitName = "Meters (user selected)"
	Else
		' User chose automatic detection
		unitFactor = 0.001 ' Default to meters (convert mm input to meters)
		unitName = "Meters (assumed)"
		
		' Try to determine units from drawing settings with multiple methods
		On Error Resume Next
		Dim insUnits As Integer
		insUnits = -1 ' Initialize to invalid value
		
		' First try: Database InsUnits
		If doc.Database Is Nothing = False Then
			insUnits = doc.Database.InsUnits
		End If
		
		' Second try: System variable INSUNITS
		If insUnits = -1 Or insUnits = 0 Then
			insUnits = doc.GetVariable("INSUNITS")
		End If
		
		' Third try: System variable LUNITS (linear units)
		If insUnits = -1 Or insUnits = 0 Then
			Dim lUnits As Integer
			lUnits = doc.GetVariable("LUNITS")
			' LUNITS doesn't directly give us the unit type, but we can make educated guess
			' If LUNITS is set, assume the drawing has been configured properly
		End If
		
		On Error GoTo 0
		
		' Interpret the units
		Select Case insUnits
			Case 1 ' Inches
				unitFactor = 25.4 ' Convert mm input to inches
				unitName = "Inches"
			Case 2 ' Feet
				unitFactor = 304.8 ' Convert mm input to feet
				unitName = "Feet"
			Case 4 ' Millimeters
				unitFactor = 1 ' No conversion needed
				unitName = "Millimeters"
			Case 5 ' Centimeters
				unitFactor = 0.1 ' Convert mm input to cm
				unitName = "Centimeters"
			Case 6 ' Meters
				unitFactor = 0.001 ' Convert mm input to meters
				unitName = "Meters"
			Case Else ' Unknown or unitless - assume meters
				unitFactor = 0.001 ' Convert mm input to meters
				unitName = "Meters (default)"
		End Select
	End If
    
	minGap = CDbl(InputBox("Enter minimum gap value (mm):", "Fix Small Gaps", "0.001")) / unitFactor
	maxGap = CDbl(InputBox("Enter maximum gap value (mm):", "Fix Small Gaps", "0.05")) / unitFactor
	
	' Debug: Show unit conversion info
	If unitChoice = vbYes Then
		MsgBox "Drawing units: " & unitName & vbCrLf & _
		       "Unit factor: " & unitFactor & vbCrLf & _
		       "Min gap in drawing units: " & minGap & vbCrLf & _
		       "Max gap in drawing units: " & maxGap, vbInformation, "Debug Info"
	Else
		MsgBox "Drawing units: " & unitName & " (Code: " & insUnits & ")" & vbCrLf & _
		       "Unit factor: " & unitFactor & vbCrLf & _
		       "Min gap in drawing units: " & minGap & vbCrLf & _
		       "Max gap in drawing units: " & maxGap, vbInformation, "Debug Info"
	End If
    
	' Collect all endpoints of lines, arcs, and polylines
	ReDim endpoints(0 To 0)
	ReDim types(0 To 0)
	idx = 0
	For Each ent In doc.ModelSpace
		Select Case ent.ObjectName
			Case "AcDbLine"
				endpoints(idx) = ent.StartPoint
				types(idx) = "L"
				idx = idx + 1
				ReDim Preserve endpoints(0 To idx)
				ReDim Preserve types(0 To idx)
				endpoints(idx) = ent.EndPoint
				types(idx) = "L"
				idx = idx + 1
				ReDim Preserve endpoints(0 To idx)
				ReDim Preserve types(0 To idx)
			Case "AcDbArc"
				endpoints(idx) = ent.StartPoint
				types(idx) = "A"
				idx = idx + 1
				ReDim Preserve endpoints(0 To idx)
				ReDim Preserve types(0 To idx)
				endpoints(idx) = ent.EndPoint
				types(idx) = "A"
				idx = idx + 1
				ReDim Preserve endpoints(0 To idx)
				ReDim Preserve types(0 To idx)
			Case "AcDbPolyline"
				Dim v As Integer
				On Error Resume Next
				Dim nVerts As Integer
				nVerts = ent.NumberOfVertices
				If Err.Number <> 0 Then
					' Not a supported polyline type, skip
					Err.Clear
				Else
					For v = 0 To nVerts - 1
						endpoints(idx) = ent.GetPointAt(v)
						types(idx) = "P"
						idx = idx + 1
						ReDim Preserve endpoints(0 To idx)
						ReDim Preserve types(0 To idx)
					Next v
				End If
				On Error GoTo 0
		End Select
	Next ent
    
	' Remove last empty slot
	If idx > 0 Then
		ReDim Preserve endpoints(0 To idx - 1)
		ReDim Preserve types(0 To idx - 1)
	End If
    
	' Search for pairs within gap range
	For i = 0 To UBound(endpoints) - 1
		For j = i + 1 To UBound(endpoints)
			pt1 = Array(endpoints(i)(0), endpoints(i)(1), endpoints(i)(2))
			pt2 = Array(endpoints(j)(0), endpoints(j)(1), endpoints(j)(2))
			dist = Sqr((pt1(0) - pt2(0)) ^ 2 + (pt1(1) - pt2(1)) ^ 2 + (pt1(2) - pt2(2)) ^ 2)
			' Debug: Show distance to verify units
			Debug.Print "Distance found: " & Format(dist, "0.000000") & " drawing units (" & Format(dist * unitFactor, "0.000") & "mm), (min: " & Format(minGap * unitFactor, "0.000") & "mm, max: " & Format(maxGap * unitFactor, "0.000") & "mm)"
			If dist >= minGap And dist <= maxGap Then
				' Zoom to region
				Call ZoomWindow(pt1, pt2)
				resp = MsgBox("Gap of " & Format(dist * unitFactor, "0.000") & "mm found. Fix this gap?", vbYesNo + vbQuestion, "Fix Small Gaps")
				If resp = vbYes Then
					midPt(0) = (pt1(0) + pt2(0)) / 2
					midPt(1) = (pt1(1) + pt2(1)) / 2
					midPt(2) = (pt1(2) + pt2(2)) / 2
					' Move both endpoints to midpoint
					Call MoveEndpointTo(i, midPt, doc, endpoints, types)
					Call MoveEndpointTo(j, midPt, doc, endpoints, types)
				End If
			End If
		Next j
	Next i
	MsgBox "Done!"
End Sub

' Helper: Zooms to a window around two points
Sub ZoomWindow(ByVal ptA, ByVal ptB)
	Dim minX As Double, minY As Double, maxX As Double, maxY As Double
	minX = Min(CDbl(ptA(0)), CDbl(ptB(0)))
	minY = Min(CDbl(ptA(1)), CDbl(ptB(1)))
	maxX = Max(CDbl(ptA(0)), CDbl(ptB(0)))
	maxY = Max(CDbl(ptA(1)), CDbl(ptB(1)))
	' Create zoom window with some padding around the gap
	Dim padding As Double
	padding = Max((maxX - minX), (maxY - minY)) * 2
	If padding < 0.1 Then padding = 0.1 ' Minimum zoom area
	
	Dim pt1(0 To 2) As Double, pt2(0 To 2) As Double
	pt1(0) = (minX + maxX) / 2 - padding / 2
	pt1(1) = (minY + maxY) / 2 - padding / 2
	pt1(2) = 0
	pt2(0) = (minX + maxX) / 2 + padding / 2
	pt2(1) = (minY + maxY) / 2 + padding / 2
	pt2(2) = 0
	
	ThisDrawing.Application.ZoomWindow pt1, pt2
End Sub

' Helper: Returns the minimum of two values
Function Min(a As Double, b As Double) As Double
	If a < b Then
		Min = a
	Else
		Min = b
	End If
End Function

' Helper: Returns the maximum of two values

Function Max(a As Double, b As Double) As Double
	If a > b Then
		Max = a
	Else
		Max = b
	End If
End Function

' Helper: Move endpoint to new position
Sub MoveEndpointTo(idx As Long, newPt As Variant, doc As AcadDocument, endpoints As Variant, ByRef types() As String)
	Dim ent As AcadEntity
	Dim k As Long, v As Integer
	k = 0
	For Each ent In doc.ModelSpace
		Select Case ent.ObjectName
			Case "AcDbLine"
				If types(idx) = "L" Then
					If IsEqual(ent.StartPoint, endpoints(idx)) Then
						ent.StartPoint = newPt
						Exit Sub
					ElseIf IsEqual(ent.EndPoint, endpoints(idx)) Then
						ent.EndPoint = newPt
						Exit Sub
					End If
				End If
			Case "AcDbArc"
				If types(idx) = "A" Then
					If IsEqual(ent.StartPoint, endpoints(idx)) Then
						ent.StartPoint = newPt
						Exit Sub
					ElseIf IsEqual(ent.EndPoint, endpoints(idx)) Then
						ent.EndPoint = newPt
						Exit Sub
					End If
				End If
			Case "AcDbPolyline"
				If types(idx) = "P" Then
					For v = 0 To ent.NumberOfVertices - 1
						If IsEqual(ent.GetPointAt(v), endpoints(idx)) Then
							ent.SetPointAt v, newPt
							Exit Sub
						End If
					Next v
				End If
		End Select
		k = k + 1
	Next ent
End Sub

' Helper: Checks if two points are equal (within tolerance)
Function IsEqual(ptA As Variant, ptB As Variant, Optional tol As Double = 0.00001) As Boolean
	IsEqual = (Abs(ptA(0) - ptB(0)) < tol) And (Abs(ptA(1) - ptB(1)) < tol) And (Abs(ptA(2) - ptB(2)) < tol)
End Function
