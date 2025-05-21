Option Explicit

Public Sub TextSorter_IndexTexts()
    Dim acadDoc As Object ' AcadDocument
    Dim modelSpace As Object ' AcadModelSpace
    Dim entity As Object ' AcadEntity
    Dim textObj As Object ' AcadText
    Dim layerName As String
    Dim ignoreValue As String
    Dim textObjects As Collection
    Dim textData() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim indexLayer As Object
    Dim foundIndexLayer As Boolean
    Dim indexText As Object
    Dim indexValue As String
    Dim indexHeight As Double
    Dim indexPoint(0 To 2) As Double
    Dim offsetX As Double

    ' 1. Get layer and ignore value from form
    On Error Resume Next
    layerName = formPerfisul01.cbCamada01.Value
    ignoreValue = formPerfisul01.cbTexto01.Value
    If Err.Number <> 0 Or layerName = "" Or layerName = "-- Selecione uma camada --" Then
        MsgBox "Selecione uma camada válida no formulário.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    Set acadDoc = ThisDrawing
    Set modelSpace = acadDoc.ModelSpace
    Set textObjects = New Collection

    ' 2. Collect all text objects in the layer except those to ignore
    For Each entity In modelSpace
        If entity.Layer = layerName Then
            If TypeOf entity Is AcadText Then
                If entity.TextString <> ignoreValue Then
                    textObjects.Add entity
                End If
            End If
        End If
    Next entity

    If textObjects.Count = 0 Then
        MsgBox "Nenhum texto encontrado na camada selecionada.", vbInformation
        Exit Sub
    End If

    ' 3. Store text and reference for sorting
    ReDim textData(1 To textObjects.Count, 1 To 2)
    For i = 1 To textObjects.Count
        Set textObj = textObjects(i)
        textData(i, 1) = textObj.TextString
        Set textData(i, 2) = textObj
    Next i

    ' 4. Sort textData alphabetically by text string (simple bubble sort)
    For i = 1 To UBound(textData, 1) - 1
        For j = i + 1 To UBound(textData, 1)
            If StrComp(textData(i, 1), textData(j, 1), vbTextCompare) > 0 Then
                temp = textData(i, 1): textData(i, 1) = textData(j, 1): textData(j, 1) = temp
                Set temp = textData(i, 2): Set textData(i, 2) = textData(j, 2): Set textData(j, 2) = temp
            End If
        Next j
    Next i

    ' 5. Create/find layer "Index" and set color to red (1)
    foundIndexLayer = False
    For Each indexLayer In acadDoc.Layers
        If indexLayer.Name = "Index" Then
            foundIndexLayer = True
            Exit For
        End If
    Next indexLayer
    If Not foundIndexLayer Then
        Set indexLayer = acadDoc.Layers.Add("Index")
    End If
    indexLayer.Color = 1 ' Red

    ' 6. Add index text to the left of each sorted text object
    For i = 1 To UBound(textData, 1)
        Set textObj = textData(i, 2)
        indexValue = CStr(i)
        indexHeight = 8 * textObj.Height
        offsetX = -20 * textObj.Height ' 12x height to the left
        Dim insPt As Variant
        insPt = textObj.InsertionPoint
        indexPoint(0) = insPt(0) + offsetX
        indexPoint(1) = insPt(1)
        indexPoint(2) = insPt(2)
        Set indexText = modelSpace.AddText(indexValue, indexPoint, indexHeight)
        indexText.Layer = "Index"
        indexText.Color = 1 ' Red
        indexText.Update
    Next i

    acadDoc.Regen acAllViewports
    MsgBox "Indexação concluída com sucesso.", vbInformation
End Sub
