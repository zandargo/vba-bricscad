Option Explicit

Dim i As Long


Public Sub TextCopier_Add(Optional ByVal layerNameParam As String = "")
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

    ' --- Form Access and Input Validation ---
    If layerNameParam = "" Then
        On Error Resume Next
        layerName = formPerfisul01.cbCamada01.Value
        If Err.Number <> 0 Then
            MsgBox "Erro ao acessar o controle de camada no formulário.", vbCritical, "Erro de Controle"
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
        If layerName = "-- Selecione uma camada --" Or layerName = "" Then
            MsgBox "Por favor, selecione uma camada válida.", vbExclamation, "Seleção de Camada"
            Exit Sub
        End If
    Else
        layerName = layerNameParam
    End If

    ' Calculate copy distance based on the height of the source text object (1.85 * height)
    ' Will be set inside the loop for each text object

    On Error GoTo ErrorHandler

    ' Get the current document and model space
    Set acadDoc = ThisDrawing
    Set modelSpace = acadDoc.ModelSpace

    ' --- Call MText_To_Text to handle MText conversion and explosion ---
    Call MText_To_Text(layerName)

    ' --- The explosion loop and manual MText conversion are no longer needed here,
    ' --- as MText_To_Text handles these operations.

    ' Regenerate to reflect any changes made by MText_To_Text
    acadDoc.Regen acAllViewports

    ' --- 3. Search all text objects in layer 'Gravação' ---
    Set textObjectsOnLayer = New Collection
    For Each entity In modelSpace ' Re-iterate modelspace as new text entities might exist after MText_To_Text
        If entity.Layer = layerName Then
            If TypeOf entity Is AcadText Then
                textObjectsOnLayer.Add entity
            End If
        End If
    Next entity    ' --- 4. If there is none, end script. Else, get text from form control ---
    If textObjectsOnLayer.Count = 0 Then
        ' MsgBox "Nenhum objeto de texto encontrado na camada '" & layerName & "' após a conversão.", vbInformation
        GoTo Cleanup
    End If


    ' Get the text from the form control cbTexto01
    On Error Resume Next
    userInputText = formPerfisul01.cbTexto01.Value
    If Err.Number <> 0 Then
        MsgBox "Erro ao acessar o controle de texto no formulário.", vbCritical, "Erro de Controle"
        Err.Clear
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    If Trim(userInputText) = "" Then
        MsgBox "Por favor, digite o texto a ser adicionado no campo 'Texto'.", vbExclamation, "Texto não informado"
        GoTo Cleanup
    End If

    ' --- 5. Copy all text objects 8mm 'down'. ---
    ' --- 6. Modify the content of copied text to the inputed value ---
    For Each textObj In textObjectsOnLayer
        Set entity = textObj ' To be explicit that textObj is an AcadEntity

        ' Calculate copy distance based on the height of the source text object
        copyDistance = 1.85 * entity.Height

        ' Use the anchor point that matches the justification so the copied text moves relative
        ' to the same reference (fixes Bottom Center staying in place).
        If entity.Alignment = acAlignmentLeft Then
            originalPoint = entity.InsertionPoint
        Else
            originalPoint = entity.TextAlignmentPoint
        End If
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
        ' For non-left justifications (e.g., Bottom Center) AutoCAD uses TextAlignmentPoint
        ' to anchor the text; setting both ensures the copy moves the expected 8mm down.
        copiedTextObj.InsertionPoint = newPoint
        copiedTextObj.TextAlignmentPoint = newPoint
        copiedTextObj.TextString = userInputText

        ' Ensure the copied text is on the correct layer (Copy method should preserve it, but good to be sure)
        If copiedTextObj.Layer <> layerName Then
            copiedTextObj.Layer = layerName
        End If
        copiedTextObj.Update ' Refresh the copied entity
    Next textObj

    acadDoc.Regen acAllViewports
    MsgBox "Objetos de texto copiados e modificados com sucesso.", vbInformation

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
    MsgBox "Ocorreu um erro: " & Err.Description & vbCrLf & _
           "Número do erro: " & Err.Number & vbCrLf & _
           "Fonte do erro: " & Err.Source, vbCritical, "Erro de Script"
    Resume Cleanup
End Sub


Public Sub TextCopier_Del(Optional ByVal layerNameParam As String = "")
    Dim acadDoc As Object ' AcadDocument
    Dim modelSpace As Object ' AcadModelSpace
    Dim entity As Object ' AcadEntity
    Dim textObj As Object ' AcadText
    Dim layerName As String
    Dim objectsOnLayer As Collection
    Dim textObjectsOnLayer As Collection
    Dim textToDeleteValue As String
    Dim objToExplode As Object
    Dim explodedItems As Variant
    Dim deletedCount As Long
    deletedCount = 0

    ' --- Form Access and Input Validation ---
    ' Directly reference the form controls (formPerfisul01 must be loaded)
    If layerNameParam = "" Then
        On Error Resume Next
        layerName = formPerfisul01.cbCamada01.Value
        If Err.Number <> 0 Then
            MsgBox "Erro ao acessar o controle de camada no formulário.", vbCritical, "Erro de Controle"
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
        If layerName = "-- Selecione uma camada --" Or layerName = "" Then
            MsgBox "Por favor, selecione uma camada válida no formulário.", vbExclamation, "Seleção de Camada Inválida"
            Exit Sub
        End If
    Else
        layerName = layerNameParam
    End If

    On Error Resume Next
    textToDeleteValue = formPerfisul01.cbTexto01.Value
    If Err.Number <> 0 Then
        MsgBox "Erro ao acessar o controle de texto no formulário.", vbCritical, "Erro de Controle"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    If Trim(textToDeleteValue) = "" Then
        MsgBox "Por favor, digite o texto a ser excluído no campo 'Texto' do formulário.", vbExclamation, "Texto para exclusão não informado"
        Exit Sub
    End If

    ' --- Main Logic ---
    On Error GoTo ErrorHandler ' Set main error handler for the rest of the sub

    ' Get the current document and model space
    Set acadDoc = ThisDrawing
    Set modelSpace = acadDoc.ModelSpace

    ' --- 1. Search for all objects in the specified layer ---
    Set objectsOnLayer = New Collection
    For Each entity In modelSpace
        If entity.Layer = layerName Then
            objectsOnLayer.Add entity
        End If
    Next entity

    ' --- 2. If there is none, end script. Else, explode them ---
    If objectsOnLayer.Count = 0 Then
        MsgBox "Nenhum objeto encontrado na camada '" & layerName & "'.", vbInformation
        GoTo Cleanup
    End If

    ' Explode collected objects (iterating 4 times, similar to TextCopier_Add)
    Dim explosionCycle As Long
    For explosionCycle = 1 To 4 ' Corrected loop syntax
        ' Re-collect objects on the layer in each cycle as new ones might appear from explosions
        Set objectsOnLayer = New Collection
        For Each entity In modelSpace
            If entity.Layer = layerName Then
                objectsOnLayer.Add entity
            End If
        Next entity
        
        If objectsOnLayer.Count = 0 Then Exit For ' No more objects to explode on this layer

        For Each objToExplode In objectsOnLayer
            On Error Resume Next ' To skip objects that cannot be exploded
            explodedItems = objToExplode.Explode
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo ErrorHandler ' Restore main error handler
        Next objToExplode
    Next explosionCycle
    
    ' Regenerate to reflect explosions before searching for text
    acadDoc.Regen acAllViewports

    ' --- 3. Search all text objects in the specified layer ---
    Set textObjectsOnLayer = New Collection
    For Each entity In modelSpace ' Re-iterate modelspace as new text entities might exist after explosion
        If entity.Layer = layerName Then
            If TypeOf entity Is AcadText Then
                textObjectsOnLayer.Add entity
            End If
        End If
    Next entity

    ' --- 4. If there are no text objects, end script. Else, get text to delete from form control ---
    If textObjectsOnLayer.Count = 0 Then
        MsgBox "Nenhum objeto de texto encontrado na camada '" & layerName & "' após a explosão.", vbInformation
        GoTo Cleanup
    End If

    ' --- 5. Delete text objects with value equal to textToDeleteValue ---
    For Each textObj In textObjectsOnLayer
        On Error Resume Next ' In case of issues with a specific text object
        If textObj.TextString = textToDeleteValue Then
            textObj.Delete
            If Err.Number = 0 Then ' Check if delete was successful
                deletedCount = deletedCount + 1
            Else
                ' Optionally log or inform about a text object that couldn't be deleted
                Err.Clear
            End If
        End If
        On Error GoTo ErrorHandler ' Restore main error handler
    Next textObj

    acadDoc.Regen acAllViewports
    MsgBox deletedCount & " objeto(s) de texto correspondente(s) a '" & textToDeleteValue & "' excluído(s) com sucesso da camada '" & layerName & "'.", vbInformation

Cleanup:
    Set acadDoc = Nothing
    Set modelSpace = Nothing
    Set objectsOnLayer = Nothing
    Set textObjectsOnLayer = Nothing
    Set entity = Nothing
    Set textObj = Nothing
    Set objToExplode = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro em TextCopier_Del: " & Err.Description & vbCrLf & _
           "Número do erro: " & Err.Number & vbCrLf & _
           "Fonte do erro: " & Err.Source, vbCritical, "Erro de Script"
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
