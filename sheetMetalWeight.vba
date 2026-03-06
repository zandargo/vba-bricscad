Option Explicit

' Calculate the total weight of sheet metal parts represented as region objects
' on the "Shapes" layer.
'
' Algorithm:
'   1. Collect all AcadRegion objects on the "Shapes" layer.
'   2. Identify outer regions — inner regions (holes/cuts) are excluded by
'      testing whether a region is fully contained inside another region.
'   3. Sum the areas of all outer regions (drawing units assumed to be mm,
'      so areas are in mm²).
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
        MsgBox "Nenhuma região encontrada na camada ""Shapes"".", _
               vbExclamation, "Peso da Chapa"
        Exit Sub
    End If

    ReDim Preserve allRegs(0 To regCount - 1)

    ' ------------------------------------------------------------------ '
    ' 3. Identify outer regions (exclude inner/hole regions)
    ' ------------------------------------------------------------------ '
    ' Sort by area descending so larger (outer) regions come first.
    Dim j As Long
    Dim swapped As Boolean
    Dim tmpReg As AcadRegion
    Do
        swapped = False
        For i = 0 To regCount - 2
            If allRegs(i).Area < allRegs(i + 1).Area Then
                Set tmpReg = allRegs(i)
                Set allRegs(i) = allRegs(i + 1)
                Set allRegs(i + 1) = tmpReg
                swapped = True
            End If
        Next i
    Loop While swapped

    ' keepFlags(i) = True  → outer region; False → contained inside another
    Dim keepFlags() As Boolean
    ReDim keepFlags(0 To regCount - 1)
    For i = 0 To regCount - 1
        keepFlags(i) = True
    Next i

    ' For each pair, test whether the smaller region is fully inside the larger.
    ' A copy-based Boolean intersection is used so the originals are untouched.
    For i = 0 To regCount - 1
        If keepFlags(i) Then
            For j = 0 To regCount - 1
                If i <> j And keepFlags(j) Then
                    ' Only check when containerReg (j) has area >= testReg (i)
                    If allRegs(j).Area >= allRegs(i).Area Then
                        Dim copyA As AcadRegion
                        Dim copyB As AcadRegion
                        On Error Resume Next
                        Set copyA = allRegs(i).Copy
                        Set copyB = allRegs(j).Copy
                        copyA.Boolean acIntersection, copyB

                        If Err.Number = 0 Then
                            ' If intersection == testReg area, testReg is fully inside containerReg
                            If Abs(copyA.Area - allRegs(i).Area) < 0.0001 Then
                                keepFlags(i) = False
                                copyA.Delete
                                If Not copyB Is Nothing Then copyB.Delete
                                Err.Clear
                                Exit For        ' No need to check other containers
                            End If
                        End If

                        ' Clean up copies whether or not an error occurred
                        If Not copyA Is Nothing Then
                            On Error Resume Next
                            copyA.Delete
                            Err.Clear
                        End If
                        If Not copyB Is Nothing Then
                            On Error Resume Next
                            copyB.Delete
                            Err.Clear
                        End If
                        On Error GoTo ErrHandler
                    End If
                End If
            Next j
        End If
    Next i

    ' ------------------------------------------------------------------ '
    ' 4. Sum areas of outer regions
    ' ------------------------------------------------------------------ '
    Dim totalAreaMm2 As Double
    totalAreaMm2 = 0
    For i = 0 To regCount - 1
        If keepFlags(i) Then
            totalAreaMm2 = totalAreaMm2 + allRegs(i).Area
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
    ' Count outer regions
    Dim outerCount As Long
    outerCount = 0
    For i = 0 To regCount - 1
        If keepFlags(i) Then outerCount = outerCount + 1
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
           "Área total (regiões externas): " & Format(totalAreaMm2, "#,##0.00") & " mm²" & vbCr & _
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
