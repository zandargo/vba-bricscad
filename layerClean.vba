Public Sub LayerClean_Del(Optional ByVal layerNameParam As String = "")
    Dim acadDoc As Object ' AcadDocument
    Dim modelSpace As Object ' AcadModelSpace
    Dim entity As Object ' AcadEntity
    Dim layerName As String
    Dim objectsOnLayer As Collection
    Dim objToDelete As Object
    Dim deletedCount As Long
    Dim i As Long

    deletedCount = 0

    ' --- Get layer name from form or parameter ---
    If layerNameParam = "" Then
        On Error Resume Next
        layerName = formPerfisul01.cbCamada02.Value
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

    On Error GoTo ErrorHandler

    ' --- Get the current document and model space ---
    Set acadDoc = ThisDrawing
    Set modelSpace = acadDoc.ModelSpace

    ' --- Collect all objects on the specified layer ---
    Set objectsOnLayer = New Collection
    For Each entity In modelSpace
        If entity.Layer = layerName Then
            objectsOnLayer.Add entity
        End If
    Next entity

    If objectsOnLayer.Count = 0 Then
        MsgBox "No objects found on layer '" & layerName & "'.", vbInformation
        GoTo Cleanup
    End If

    ' --- Delete all objects on the layer ---
    For i = objectsOnLayer.Count To 1 Step -1
        Set objToDelete = objectsOnLayer(i)
        On Error Resume Next
        objToDelete.Delete
        If Err.Number = 0 Then
            deletedCount = deletedCount + 1
        Else
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next i

    acadDoc.Regen acAllViewports
    MsgBox deletedCount & " object(s) deleted from layer '" & layerName & "'.", vbInformation

Cleanup:
    Set acadDoc = Nothing
    Set modelSpace = Nothing
    Set objectsOnLayer = Nothing
    Set entity = Nothing
    Set objToDelete = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in LayerClean_Del: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: " & Err.Source, vbCritical, "Script Error"
    Resume Cleanup
End Sub
