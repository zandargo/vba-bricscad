Option Explicit

Public Sub FixDuplicatedCodes()
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
    Dim dict As Object
    Dim txt As String
    Dim changedCount As Long
    Dim suffix As String

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

    ' 5. First pass: count occurrences of each text
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(textData, 1)
        txt = textData(i, 1)
        If dict.Exists(txt) Then
            dict(txt) = dict(txt) + 1
        Else
            dict(txt) = 1
        End If
    Next i
    
    ' 6. Second pass: add suffixes to duplicated texts
    Dim currentCount As Object
    Set currentCount = CreateObject("Scripting.Dictionary")
    changedCount = 0

    For i = 1 To UBound(textData, 1)
        Set textObj = textData(i, 2)
        txt = textData(i, 1)
        
        If dict(txt) > 1 Then
            ' This text has duplicates, add suffix
            If currentCount.Exists(txt) Then
                currentCount(txt) = currentCount(txt) + 1
            Else
                currentCount(txt) = 1
            End If
            suffix = Right("0" & CStr(currentCount(txt)), 2) ' Format as 01, 02, etc.
            textObj.TextString = txt & suffix
            textObj.Update
            changedCount = changedCount + 1
        End If
    Next i

    acadDoc.Regen acAllViewports

    ' Display result message
    MsgBox "Correção de códigos duplicados concluída." & vbCrLf & _
           "Quantidade de valores alterados: " & changedCount, vbInformation
End Sub
