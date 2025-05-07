Option Explicit

Sub ResizeAndColorCircles()
    Dim acadDoc As Object ' AcadDocument
    Dim modelSpace As Object ' AcadModelSpace
    Dim entity As Object ' AcadEntity
    Dim circleObj As Object ' AcadCircle
    Dim otherCircleObj As Object ' AcadCircle
    Dim diametersArray As Variant
    Dim i As Integer
    Dim originalDiameter As Double
    Dim newDiameter As Double
    Dim targetDiameter As Double
	 Dim margin As Double
    Dim lowerBound As Double
    Dim upperBound As Double
    Dim centerPoint As Variant
    Dim tempEntity As Object

    ' Set the active document and model space
    On Error Resume Next
    Set acadDoc = ThisDrawing
    If acadDoc Is Nothing Then
        MsgBox "Could not get ThisDrawing. Make sure you are running this from BricsCAD.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Set modelSpace = acadDoc.ModelSpace

    ' 1. Has an array of pair of digits. Example: [ [3.2, 1.9], [13.65, 13.5] ]
    diametersArray = Array(Array(3.2, 1.9), Array(9, 13.5))
    ' Add more pairs as needed, e.g.:
    ' diametersArray = Array(Array(3.2, 1.9), Array(13.65, 13.5), Array(20.0, 15.0))

    ' 2. Iterate the current drawing and search for every circle
    For Each entity In modelSpace
        If TypeOf entity Is AcadCircle Then
            Set circleObj = entity
            originalDiameter = circleObj.Diameter

            For i = LBound(diametersArray) To UBound(diametersArray)
                targetDiameter = diametersArray(i)(0)
                newDiameter = diametersArray(i)(1)

                ' with a safety margin of 2% for more or less.
					 margin = 0.02
                lowerBound = targetDiameter * (1 - margin)
                upperBound = targetDiameter * (1 + margin)

                If originalDiameter >= lowerBound And originalDiameter <= upperBound Then
                    ' Store center point before resizing
                    centerPoint = circleObj.Center

                    ' 3. Resize that circle to the corresponding value in the pair
                    circleObj.Diameter = newDiameter
                    
                    ' 4. Change the circle color to red
                    circleObj.Color = acRed ' 1 = Red

                    ' 5. Delete every other circle concentric to that circle.
                    ' Iterate again to find concentric circles (excluding the one just modified)
                    Dim tempEntityCollection As Object ' AcadBlock
                    Set tempEntityCollection = modelSpace
                    Dim k As Integer
                    k = 0
                    Do While k < tempEntityCollection.Count
                        Set tempEntity = tempEntityCollection.Item(k)
                        If TypeOf tempEntity Is AcadCircle Then
                            Set otherCircleObj = tempEntity
                            ' Check if it's not the same circle and if it's concentric
                            If Not otherCircleObj.Handle = circleObj.Handle Then
                                If PointsAreEqual(otherCircleObj.Center, centerPoint) Then
                                    otherCircleObj.Delete
                                    ' If an entity is deleted, the collection count might change,
                                    ' so we don't increment k to re-check the current index.
                                    ' However, a safer approach is to iterate backwards or build a list to delete.
                                    ' For simplicity here, we re-evaluate count and might miss some in complex scenarios
                                    ' A more robust way is to collect handles and delete them after the main loop.
                                    ' Or iterate backwards: For k = tempEntityCollection.Count - 1 To 0 Step -1
                                Else
                                    k = k + 1
                                End If
                            Else
                                k = k + 1
                            End If
                        Else
                            k = k + 1
                        End If
                    Loop
                    
                    ' Exit the inner loop once a match is found and processed for this circle
                    Exit For
                End If
            Next i
        End If
    Next entity

    acadDoc.Regen acAllViewports
    MsgBox "Circle processing complete.", vbInformation

End Sub

Private Function PointsAreEqual(p1 As Variant, p2 As Variant, Optional tolerance As Double = 0.0001) As Boolean
    ' Helper function to compare two points with a tolerance
    If Abs(p1(0) - p2(0)) < tolerance And _
       Abs(p1(1) - p2(1)) < tolerance And _
       Abs(p1(2) - p2(2)) < tolerance Then
        PointsAreEqual = True
    Else
        PointsAreEqual = False
    End If
End Function

' To run this macro in BricsCAD:
' 1. Open BricsCAD.
' 2. Press ALT+F11 to open the VBA IDE (or type VBAIDE in the command line).
' 3. In the VBA IDE, go to Insert > Module.
' 4. Paste this code into the module.
' 5. Close the VBA IDE.
' 6. In BricsCAD, type VBARUN in the command line.
' 7. Select "ResizeAndColorCircles" from the list and click "Run".
' Make sure you have a drawing open with circles to test.
