Option Explicit

Sub TextCopier_Main()
    Dim acadDoc As Object ' AcadDocument
    Dim modelSpace As Object ' AcadModelSpace
    Dim entity As Object ' AcadEntity
    Dim textObj As Object ' AcadText
    Dim layerName As String
    Dim objectsOnLayer As Collection
    Dim textObjectsOnLayer As Collection
    Dim i As Long
    Dim userInputText As String
    Dim copyDistance As Double
    Dim originalPoint As Variant
    Dim newPoint As Variant
    Dim rotationAngle As Double
    Dim dx As Double
    Dim dy As Double
    Dim copiedTextObj As Object
    Dim objToExplode As Object
    Dim explodedItems As Variant

    layerName = "Gravação"
    copyDistance = 8.0 ' 8mm

    On Error GoTo ErrorHandler

    ' Get the current document and model space
    Set acadDoc = ThisDrawing
    Set modelSpace = acadDoc.ModelSpace

    ' --- 1. Search for all objects in layer 'Gravação' ---
    Set objectsOnLayer = New Collection
    For Each entity In modelSpace
        If entity.Layer = layerName Then
            objectsOnLayer.Add entity
        End If
    Next entity

    ' --- 2. If there is none, end script. Else, explode them ---
    If objectsOnLayer.Count = 0 Then
        MsgBox "No objects found on layer '" & layerName & "'.", vbInformation
        GoTo Cleanup
    End If

    ' Explode collected objects
    ' Iterating a separate collection for explosion is safer
    For Each objToExplode In objectsOnLayer
        On Error Resume Next ' To skip objects that cannot be exploded
        explodedItems = objToExplode.Explode
        ' objToExplode is now deleted if successfully exploded.
        ' explodedItems contains an array of new entities.
        ' We don't need to process explodedItems directly here,
        ' as the next step will re-scan the modelspace for texts.
        If Err.Number <> 0 Then
            ' Optional: Log or notify if an object couldn't be exploded
            ' MsgBox "Could not explode object: " & objToExplode.ObjectName & " (Handle: " & objToExplode.Handle & ")"
            Err.Clear
        End If
        On Error GoTo ErrorHandler ' Restore main error handler
    Next objToExplode
    
    ' Regenerate to reflect explosions before searching for text
    acadDoc.Regen acAllViewports

    ' --- 3. Search all text objects in layer 'Gravação' ---
    Set textObjectsOnLayer = New Collection
    For Each entity In modelSpace ' Re-iterate modelspace as new text entities might exist after explosion
        If entity.Layer = layerName Then
            If TypeOf entity Is AcadText Then
                textObjectsOnLayer.Add entity
            End If
        End If
    Next entity

    ' --- 4. If there is none, end script. Else, ask the user the name of text to be added. ---
    If textObjectsOnLayer.Count = 0 Then
        MsgBox "No text objects found on layer '" & layerName & "' after explosion.", vbInformation
        GoTo Cleanup
    End If

    userInputText = InputBox("Enter the text to be added to the copied texts:", "Text Input", "Carroceria") ' Added example as default
    If userInputText = "" Then
        MsgBox "No text input provided. Script will exit.", vbInformation
        GoTo Cleanup
    End If

    ' --- 5. Copy all text objects 8mm 'down'. ---
    ' --- 6. Modify the content of copied text to the inputed value ---
    For Each textObj In textObjectsOnLayer
        Set entity = textObj ' To be explicit that textObj is an AcadEntity
        
        originalPoint = entity.InsertionPoint
        rotationAngle = entity.Rotation ' This is in Radians

        ' Calculate offset based on rotation
        ' "Down" relative to text: dx = distance * sin(angle), dy = distance * -cos(angle)
        dx = copyDistance * Sin(rotationAngle)
        dy = copyDistance * (-Cos(rotationAngle))

        ReDim newPoint(0 To 2) As Double
        newPoint(0) = originalPoint(0) + dx
        newPoint(1) = originalPoint(1) + dy
        newPoint(2) = originalPoint(2) ' Preserve Z coordinate

        ' Create a copy of the text object
        Set copiedTextObj = entity.Copy

        ' Modify the copied text
        copiedTextObj.InsertionPoint = newPoint
        copiedTextObj.TextString = userInputText
        
        ' Ensure the copied text is on the correct layer (Copy method should preserve it, but good to be sure)
        If copiedTextObj.Layer <> layerName Then
            copiedTextObj.Layer = layerName
        End If
        copiedTextObj.Update ' Refresh the copied entity
    Next textObj

    acadDoc.Regen acAllViewports
    MsgBox "Text objects copied and modified successfully.", vbInformation

Cleanup:
    Set acadDoc = Nothing
    Set modelSpace = Nothing
    Set objectsOnLayer = Nothing
    Set textObjectsOnLayer = Nothing
    Set entity = Nothing
    Set textObj = Nothing
    Set copiedTextObj = Nothing
    Set objToExplode = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: " & Err.Source, vbCritical, "Script Error"
    Resume Cleanup
End Sub

' Instructions to run:
' 1. Open BricsCAD.
' 2. Press ALT+F11 to open the VBA IDE (or type VBAIDE in the command line).
' 3. In the VBA IDE, go to File > Import File... and select this .vba file.
'    Alternatively, go to Insert > Module and paste this code into the new module.
' 4. Close the VBA IDE.
' 5. In BricsCAD, type VBARUN in the command line.
' 6. Select "TextCopier_Main" from the list and click "Run".
' 7. Ensure your drawing has a layer named "Gravação" with some objects and text entities for testing.
