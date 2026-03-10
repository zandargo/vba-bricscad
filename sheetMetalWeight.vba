Option Explicit

' Calculate the total weight of sheet metal parts represented as region objects
' on the "Shapes" layer.
'
' Algorithm:
'   1. Collect all AcadRegion objects on the "Shapes" layer.
'   2. Compute the nesting depth of every region: how many other regions
'      fully contain it (determined via Boolean intersection copies).
'   3. Sum areas using the even-odd rule:
'        even depth (0, 2, …) → add area  (outer material / island)
'        odd  depth (1, 3, …) → subtract area (hole / cut)
'      Drawing units assumed to be mm, so areas are in mm².
'   4. Compute weight using values from formPerfisul01:
'        thickness  = valSheetThickness  (mm)
'        density    = valSheetDensity    (kg/m³)
'        weight (kg) = (totalArea_mm² / 1 000 000) × (thickness_mm / 1 000) × density_kg_m³
'   5. Insert a Text entity at (0, -40) with height 20 mm:
'        "Peso Total: XXX,XX kg"  (comma as decimal separator)

Public Sub CalculateSheetMetalWeight()
    On Error GoTo ErrHandler

    Dim doc As AcadDocument
    Set doc = ThisDrawing

    ' ------------------------------------------------------------------ '
    ' 1. Read parameters from the form
    ' ------------------------------------------------------------------ '
    Dim thicknessMm As Double
    Dim densityKgM3 As Double

    On Error Resume Next
    thicknessMm = CDbl(formPerfisul01.valSheetThickness.Value)
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Não foi possível ler a espessura (valSheetThickness) do formulário.", _
               vbExclamation, "Peso da Chapa"
        Exit Sub
    End If
    densityKgM3 = CDbl(formPerfisul01.valSheetDensity.Value)
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Não foi possível ler a densidade (valSheetDensity) do formulário.", _
               vbExclamation, "Peso da Chapa"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    If thicknessMm <= 0 Then
        MsgBox "Espessura inválida: " & thicknessMm & " mm.", vbExclamation, "Peso da Chapa"
        Exit Sub
    End If
    If densityKgM3 <= 0 Then
        MsgBox "Densidade inválida: " & densityKgM3 & " kg/m³.", vbExclamation, "Peso da Chapa"
        Exit Sub
    End If

    ' ------------------------------------------------------------------ '
    ' 2. Collect all AcadRegion objects on layer "Shapes"
    ' ------------------------------------------------------------------ '
    Dim allRegs() As AcadRegion
    Dim regCount As Long
    regCount = 0
    ReDim allRegs(0 To doc.ModelSpace.Count - 1)

    Dim i As Long
    For i = 0 To doc.ModelSpace.Count - 1
        Dim ent As AcadEntity
        Set ent = doc.ModelSpace.Item(i)
        If UCase$(Trim$(ent.Layer)) = "SHAPES" Then
            If TypeOf ent Is AcadRegion Then
                Set allRegs(regCount) = ent
                regCount = regCount + 1
            End If
        End If
    Next i

    If regCount = 0 Then
        MsgBox "Nenhuma região verde encontrada na camada ""Shapes"".", _
               vbExclamation, "Peso da Chapa"
        Exit Sub
    End If

    ReDim Preserve allRegs(0 To regCount - 1)

    ' ------------------------------------------------------------------ '
    ' 3. Compute nesting depth for each region (even-odd rule)
    '    nestDepth(i) = number of other regions that fully contain region i.
    '    Even depth → material (add); Odd depth → hole/cut (subtract).
    ' ------------------------------------------------------------------ '
    Dim j As Long
    Dim nestDepth() As Long
    ReDim nestDepth(0 To regCount - 1)
    For i = 0 To regCount - 1
        nestDepth(i) = 0
    Next i

    Dim copyA As AcadRegion
    Dim copyB As AcadRegion

    ' For each region i, count how many other regions j fully contain it.
    ' "Fully contained" means intersect(i, j) has the same area as i.
    ' Copies are used for the Boolean op so originals are never touched.
    For i = 0 To regCount - 1
        For j = 0 To regCount - 1
            If i <> j Then
                If allRegs(j).Area >= allRegs(i).Area Then
                    On Error Resume Next
                    Set copyA = allRegs(i).Copy
                    Set copyB = allRegs(j).Copy
                    Err.Clear
                    copyA.Boolean acIntersection, copyB

                    If Err.Number = 0 Then
                        If Abs(copyA.Area - allRegs(i).Area) < 0.0001 Then
                            nestDepth(i) = nestDepth(i) + 1
                        End If
                    End If

                    If Not copyA Is Nothing Then copyA.Delete
                    If Not copyB Is Nothing Then copyB.Delete
                    Err.Clear
                    On Error GoTo ErrHandler
                End If
            End If
        Next j
    Next i

    ' ------------------------------------------------------------------ '
    ' 4. Sum areas using even-odd rule: even depth → add, odd → subtract
    ' ------------------------------------------------------------------ '
    Dim totalAreaMm2 As Double
    totalAreaMm2 = 0
    For i = 0 To regCount - 1
        If (nestDepth(i) Mod 2) = 0 Then
            totalAreaMm2 = totalAreaMm2 + allRegs(i).Area
        Else
            totalAreaMm2 = totalAreaMm2 - allRegs(i).Area
        End If
    Next i

    ' ------------------------------------------------------------------ '
    ' 5. Calculate total weight
    '    weight (kg) = area (m²) × thickness (m) × density (kg/m³)
    '    area_m²     = totalAreaMm2 / 1 000 000
    '    thickness_m = thicknessMm  / 1 000
    ' ------------------------------------------------------------------ '
    Dim weightKg As Double
    weightKg = (totalAreaMm2 / 1000000#) * (thicknessMm / 1000#) * densityKgM3

    ' ------------------------------------------------------------------ '
    ' 6. Compute drawing-content bounding box to derive dynamic text metrics
    '    (considering only objects in layer "linha_grade")
    ' ------------------------------------------------------------------ '
    Dim bbMinX As Double, bbMinY As Double, bbMaxX As Double, bbMaxY As Double
    Dim bbFirst As Boolean: bbFirst = True
    Dim bbEnt As AcadEntity
    For Each bbEnt In doc.ModelSpace
        If UCase$(Trim$(bbEnt.Layer)) = "0" Or UCase$(Trim$(bbEnt.Layer)) = "LINHA_GRADE" Then
            Dim bbMin As Variant, bbMax As Variant
            On Error Resume Next
            bbEnt.GetBoundingBox bbMin, bbMax
            If Err.Number = 0 Then
                If bbFirst Then
                    bbMinX = bbMin(0): bbMinY = bbMin(1)
                    bbMaxX = bbMax(0): bbMaxY = bbMax(1)
                    bbFirst = False
                Else
                    If bbMin(0) < bbMinX Then bbMinX = bbMin(0)
                    If bbMin(1) < bbMinY Then bbMinY = bbMin(1)
                    If bbMax(0) > bbMaxX Then bbMaxX = bbMax(0)
                    If bbMax(1) > bbMaxY Then bbMaxY = bbMax(1)
                End If
            End If
            Err.Clear
            On Error GoTo ErrHandler
        End If
    Next bbEnt

    Dim dX As Double, dY As Double
    If bbFirst Then
        ' Fallback: no entities found, use safe defaults
        dX = 1000
        dY = 1000
    Else
        dX = bbMaxX - bbMinX
        dY = bbMaxY - bbMinY
    End If

    ' ------------------------------------------------------------------ '
    ' 7. Insert result text using dynamic position and height
    '    textHeight = 0.015 * dY
    '    position   = ( 0.0035 * dY , 0.0035 * dY )
    ' ------------------------------------------------------------------ '
    ' Count outer (depth-0) regions = number of distinct parts
    Dim outerCount As Long
    outerCount = 0
    For i = 0 To regCount - 1
        If nestDepth(i) = 0 Then outerCount = outerCount + 1
    Next i

    ' Format with 2 decimal places, replacing "." with "," per Brazilian locale
    Dim weightStr As String
    weightStr = Format(weightKg, "0.00")
    weightStr = Replace(weightStr, ".", ",")

    Dim labelText As String
    labelText = "Qtd de peças: " & Format(outerCount, "00") & "   |   Peso Total: " & weightStr & " kg"

    Dim textHeight As Double
    textHeight = 0.01 * dY

    Dim insertPt(0 To 2) As Double
    insertPt(0) = 0.0065 * dY
    insertPt(1) = 0.0065 * dY
    insertPt(2) = 0

    doc.StartUndoMark

    ' Remove any previous weight label to avoid stacking on re-runs
    Dim existingEnt As AcadEntity
    For Each existingEnt In doc.ModelSpace
        If TypeOf existingEnt Is AcadText Then
            On Error Resume Next
            Dim txtObj As AcadText
            Set txtObj = existingEnt
            If InStr(1, txtObj.TextString, "Quantidade de peças:", vbTextCompare) > 0 Or _
               InStr(1, txtObj.TextString, "Peso Total:", vbTextCompare) > 0 Then
                txtObj.Delete
            End If
            Err.Clear
            On Error GoTo ErrHandler
        End If
    Next existingEnt

    Dim newText As AcadText
    Set newText = doc.ModelSpace.AddText(labelText, insertPt, textHeight)
    newText.Layer = "0"
    newText.Color = acByLayer

    doc.EndUndoMark

    MsgBox "Peso calculado com sucesso!" & vbCr & vbCr & _
           "Quantidade de peças: " & outerCount & vbCr & _
           "Área líquida (externas - furos): " & Format(totalAreaMm2 / 1000000#, "#,##0.000000") & " m²" & vbCr & _
           "Espessura: " & thicknessMm & " mm" & vbCr & _
           "Densidade: " & densityKgM3 & " kg/m³" & vbCr & vbCr & _
           labelText, _
           vbInformation, "Peso da Chapa"

    Exit Sub

ErrHandler:
    MsgBox "Erro ao calcular peso: " & Err.Description, vbCritical, "Peso da Chapa"
    On Error Resume Next
    doc.EndUndoMark
End Sub

' ======================================================================== '
' ToggleShapeExclusion
' -----------------------------------------------------------------------  '
' Temporarily exclude (or re-include) individual regions from the weight
' calculation without touching their color.
'
' Mechanism:
'   - Active regions live on layer "Shapes"  (green)  → counted by weight calc.
'   - Excluded regions live on layer "Shapes_Skip" (red) → ignored by weight calc.
'
' Usage: run this sub, select one or more regions; each one is toggled
'   between the two layers. Run it again on the same region to restore it.
' ======================================================================== '
Public Sub ToggleShapeExclusion()
    On Error GoTo ErrHandler

    Dim doc As AcadDocument
    Set doc = ThisDrawing

    ' Hide the form so the user can interact freely with the drawing
    On Error Resume Next
    formPerfisul01.Hide
    On Error GoTo ErrHandler

    ' Ensure both layers exist before saving state (so they appear in the snapshot)
    EnsureShapesSkipLayer doc

    ' ------------------------------------------------------------------ '
    ' Save layer visibility state, then hide everything except the two
    ' Shapes layers so the user sees only what they need to select.
    ' ------------------------------------------------------------------ '
    Dim savedCount As Long
    savedCount = 0
    Dim savedNames() As String
    Dim savedOn() As Boolean
    ReDim savedNames(0 To doc.Layers.Count - 1)
    ReDim savedOn(0 To doc.Layers.Count - 1)

    Dim lyr As AcadLayer
    For Each lyr In doc.Layers
        savedNames(savedCount) = lyr.Name
        savedOn(savedCount) = lyr.LayerOn
        savedCount = savedCount + 1
        Dim layUp As String
        layUp = UCase$(Trim$(lyr.Name))
        If layUp <> "SHAPES" And layUp <> "SHAPES_SKIP" Then
            On Error Resume Next
            lyr.LayerOn = False
            Err.Clear
            On Error GoTo ErrHandler
        Else
            On Error Resume Next
            lyr.LayerOn = True
            Err.Clear
            On Error GoTo ErrHandler
        End If
    Next lyr
    doc.Regen acAllViewports

    ' ------------------------------------------------------------------ '
    ' Build a selection set and let the user pick regions
    ' ------------------------------------------------------------------ '
    Dim ss As AcadSelectionSet
    On Error Resume Next
    Set ss = doc.SelectionSets.Item("SMW_TOGGLE")
    If Err.Number = 0 Then ss.Delete
    Err.Clear
    On Error GoTo ErrHandler
    Set ss = doc.SelectionSets.Add("SMW_TOGGLE")

    doc.Utility.Prompt vbCr & "Selecione regiões para (des)ativar no cálculo..." & vbCr
    ss.SelectOnScreen

    If ss.Count = 0 Then
        MsgBox "Nenhum objeto selecionado.", vbExclamation, "Excluir/Incluir Forma"
        GoTo Cleanup
    End If

    doc.StartUndoMark

    Dim ent As AcadEntity
    Dim toggled As Long
    toggled = 0
    For Each ent In ss
        If Not TypeOf ent Is AcadRegion Then GoTo NextEnt
        Dim entLayer As String
        entLayer = UCase$(Trim$(ent.Layer))
        Select Case entLayer
            Case "SHAPES"
                ent.Layer = "Shapes_Skip"
            Case "SHAPES_SKIP"
                ent.Layer = "Shapes"
            Case Else
                GoTo NextEnt   ' not a managed shape — skip silently
        End Select
        toggled = toggled + 1
NextEnt:
    Next ent

    doc.EndUndoMark

    If toggled = 0 Then
        MsgBox "Nenhuma região nas camadas ""Shapes"" ou ""Shapes_Skip"" foi selecionada.", _
               vbExclamation, "Excluir/Incluir Forma"
    Else
        MsgBox toggled & " região(ões) alternada(s) entre ""Shapes"" e ""Shapes_Skip""." & vbCr & vbCr & _
               "Regiões em ""Shapes_Skip"" (vermelho) serão ignoradas no cálculo de peso.", _
               vbInformation, "Excluir/Incluir Forma"
    End If

    GoTo Cleanup

ErrHandler:
    MsgBox "Erro ao alternar exclusão: " & Err.Description, vbCritical, "Excluir/Incluir Forma"
    On Error Resume Next
    doc.EndUndoMark
Cleanup:
    ' Restore every layer to its original visibility state
    On Error Resume Next
    Dim k As Long
    For k = 0 To savedCount - 1
        Dim lyrR As AcadLayer
        Set lyrR = doc.Layers.Item(savedNames(k))
        If Err.Number = 0 Then lyrR.LayerOn = savedOn(k)
        Err.Clear
    Next k
    If savedCount > 0 Then doc.Regen acAllViewports

    ss.Delete

    ' Restore the form
    formPerfisul01.Show
    On Error GoTo 0
End Sub

Private Sub EnsureShapesSkipLayer(doc As AcadDocument)
    On Error Resume Next
    Dim lyr As AcadLayer
    Set lyr = doc.Layers.Item("Shapes_Skip")
    If Err.Number <> 0 Then
        Err.Clear
        Set lyr = doc.Layers.Add("Shapes_Skip")
        lyr.Color = acRed          ' visually distinct from the green "Shapes" layer
    End If
    On Error GoTo 0
End Sub
