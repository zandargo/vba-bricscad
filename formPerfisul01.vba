Private Sub UserForm_Initialize()
    Dim acadDoc As Object ' AcadDocument
    Dim acadLayers As Object ' AcadLayers collection
    Dim layerObj As Object ' AcadLayer
    
    On Error Resume Next
    Set acadDoc = ThisDrawing
    
    If acadDoc Is Nothing Then
        MsgBox "Could not get ThisDrawing. Make sure you are running this from BricsCAD.", vbCritical
        Exit Sub
    End If
    
    Set acadLayers = acadDoc.Layers
    
    ' Clear the combo box first
    cbCamada01.Clear
    cbCamada02.Clear
    
    ' Add a default option as the first item
    cbCamada01.AddItem "-- Selecione uma camada --"
    cbCamada01.ListIndex = 0 ' Select the default option
    cbCamada02.AddItem "-- Selecione uma camada --"
    cbCamada02.ListIndex = 0 ' Select the default option
    
    ' Loop through all layers and add them to the combo box
    For Each layerObj In acadLayers
        cbCamada01.AddItem layerObj.Name
        cbCamada02.AddItem layerObj.Name
    Next layerObj

	 cbTexto01.AddItem "CARROCERIA"
	 cbTexto01.AddItem "FERRAMENTARIA"
	 cbTexto01.AddItem "PORTAS"
	 cbTexto01.AddItem "TAMPA"
	 cbTexto01.AddItem "TETO"
	 cbTexto01.AddItem "VIDROS"
    
    On Error GoTo 0
End Sub

Private Sub btnPuncionadeira_Click()
 Call ResizeAndColorCircles
End Sub

Private Sub btnAddText01_Click()
    ' Check if a layer is selected
    If cbCamada01.ListIndex <= 0 Then
        MsgBox "Por favor, selecione uma camada antes de continuar.", vbExclamation, "Seleção de Camada"
        Exit Sub
    End If
    
    ' Check if text is provided
    If Trim(cbTexto01.Value) = "" Then
        MsgBox "Por favor, informe o texto a ser adicionado.", vbExclamation, "Texto não Informado"
        cbTexto01.SetFocus
        Exit Sub
    End If
    
    ' Call TextCopier_Add with the selected layer name
    Call TextCopier_Add(cbCamada01.Value)
End Sub

Private Sub btnDelText01_Click()
    ' Check if a layer is selected
    If cbCamada01.ListIndex <= 0 Then
        MsgBox "Por favor, selecione uma camada antes de continuar.", vbExclamation, "Seleção de Camada"
        Exit Sub
    End If
    
    ' Check if text is provided
    If Trim(cbTexto01.Value) = "" Then
        MsgBox "Por favor, informe o texto a ser adicionado.", vbExclamation, "Texto não Informado"
        cbTexto01.SetFocus
        Exit Sub
    End If
    
    ' Call TextCopier_Add with the selected layer name
    Call TextCopier_Del(cbCamada01.Value)
End Sub


Private Sub btnLimparCamada01_Click()
    Call LayerClean_Del(cbCamada02.Value)
End Sub
