Option Explicit

' Auto-orient a selection by finding the best rotation angle (0-180 degrees).
' Call from form with: AutoOrientFromForm(Me)
' Call standalone with: AutoOrient()

Public Sub AutoOrient()
    AutoOrientCore Nothing
End Sub

Public Sub AutoOrientFromForm(frm As Object)
    AutoOrientCore frm
End Sub

Private Sub AutoOrientCore(frm As Object)
    Dim doc As AcadDocument
    Set doc = ThisDrawing
    
    Dim wasVisible As Boolean: wasVisible = False
    If Not frm Is Nothing Then
        On Error Resume Next
        wasVisible = frm.Visible
        frm.Hide
        On Error GoTo 0
    End If
    
    On Error GoTo ErrorHandler
    
    ' Prompt user to select entities
    Dim ss As AcadSelectionSet
    Set ss = doc.SelectionSets.Add("AUTOORIENT_SS")
    
    doc.Utility.Prompt vbCr & "Selecione os objetos: " & vbCr
    ss.SelectOnScreen
    
    If ss.Count = 0 Then
        doc.Utility.Prompt "Nenhum objeto selecionado." & vbCr
        ss.Delete
        GoTo Cleanup
    End If
    
    doc.StartUndoMark
    
    ' Get bounding box of current selection
    Dim minPt As Variant, maxPt As Variant
    GetSelectionBounds ss, minPt, maxPt
    
    Dim centerPt(0 To 2) As Double
    centerPt(0) = (minPt(0) + maxPt(0)) / 2
    centerPt(1) = (minPt(1) + maxPt(1)) / 2
    centerPt(2) = 0
    
    ' Find best rotation angle
    Dim bestAngle As Double
    Dim bestHeight As Double
    bestAngle = FindBestRotationAngle(ss, centerPt, bestHeight)
    
    Dim degAngle As Double
    degAngle = bestAngle * 180 / 3.14159265358979
    
    doc.Utility.Prompt "Melhor angulo: " & Format(degAngle, "0.00") & " graus" & vbCr
    
    ' Apply rotation if angle is significant
    If Abs(bestAngle) > 0.001 Then
        RotateSelectedEntities ss, centerPt, bestAngle
        MsgBox "Rotacionado em " & Format(degAngle, "0.00") & " graus.", vbInformation
    Else
        MsgBox "Nenhuma rotacao necessaria (angulo ~0).", vbInformation
    End If
    
    doc.EndUndoMark
    ss.Delete
    
Cleanup:
    If Not frm Is Nothing And wasVisible Then
        On Error Resume Next
        frm.Show
        On Error GoTo 0
    End If
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not ss Is Nothing Then ss.Delete
    MsgBox "Erro: " & Err.Description, vbCritical
    If Not frm Is Nothing And wasVisible Then frm.Show
End Sub

' Get bounding box of all entities in selection set
Private Sub GetSelectionBounds(ss As AcadSelectionSet, ByRef minPt As Variant, ByRef maxPt As Variant)
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double
    Dim first As Boolean: first = True
    
    Dim ent As AcadEntity
    Dim entMinPt As Variant, entMaxPt As Variant
    
    For Each ent In ss
        On Error Resume Next
        ent.GetBoundingBox entMinPt, entMaxPt
        
        If Not first Then
            If entMinPt(0) < minX Then minX = entMinPt(0)
            If entMinPt(1) < minY Then minY = entMinPt(1)
            If entMaxPt(0) > maxX Then maxX = entMaxPt(0)
            If entMaxPt(1) > maxY Then maxY = entMaxPt(1)
        Else
            minX = entMinPt(0)
            minY = entMinPt(1)
            maxX = entMaxPt(0)
            maxY = entMaxPt(1)
            first = False
        End If
        On Error GoTo 0
    Next ent
    
    minPt = Array(minX, minY, 0)
    maxPt = Array(maxX, maxY, 0)
End Sub

' Find best rotation angle by testing increments from 0 to 180 degrees
Private Function FindBestRotationAngle(ss As AcadSelectionSet, centerPt() As Double, ByRef heightOut As Double) As Double
    Const STEP_DEG As Double = 2
    Const PI As Double = 3.14159265358979
    
    Dim bestAngle As Double: bestAngle = 0
    Dim bestHeight As Double: bestHeight = 1E+30
    Dim bestAspect As Double: bestAspect = 0
    
    Dim angle As Double
    Dim deg As Double
    Dim width As Double, height As Double, aspect As Double
    Dim minPt As Variant, maxPt As Variant
    
    ' Test angles from 0 to 180 degrees
    For deg = 0 To 180 Step STEP_DEG
        angle = deg * PI / 180
        
        ' Get bounds if entities were rotated by this angle
        GetRotatedBounds ss, centerPt, angle, minPt, maxPt
        
        width = maxPt(0) - minPt(0)
        height = maxPt(1) - minPt(1)
        
        If height > 0 Then
            aspect = width / height
        Else
            aspect = 0
        End If
        
        ' Prefer smaller height; use aspect ratio as tiebreaker
        If height < bestHeight - 0.001 Or (Abs(height - bestHeight) < 0.001 And aspect > bestAspect) Then
            bestHeight = height
            bestAspect = aspect
            bestAngle = angle
        End If
    Next deg
    
    heightOut = bestHeight
    FindBestRotationAngle = bestAngle
End Function

' Compute rotated bounding box for all entities at given angle
Private Sub GetRotatedBounds(ss As AcadSelectionSet, centerPt() As Double, angle As Double, _
    ByRef minPt As Variant, ByRef maxPt As Variant)
    
    Const LARGE As Double = 1E+30
    Dim minX As Double: minX = LARGE
    Dim minY As Double: minY = LARGE
    Dim maxX As Double: maxX = -LARGE
    Dim maxY As Double: maxY = -LARGE
    
    Dim cosA As Double: cosA = Cos(angle)
    Dim sinA As Double: sinA = Sin(angle)
    
    Dim ent As AcadEntity
    Dim entMinPt As Variant, entMaxPt As Variant
    Dim corners(0 To 3) As Variant
    
    For Each ent In ss
        On Error Resume Next
        ent.GetBoundingBox entMinPt, entMaxPt
        
        ' Get 4 corners of entity bounding box
        corners(0) = Array(entMinPt(0), entMinPt(1))
        corners(1) = Array(entMaxPt(0), entMinPt(1))
        corners(2) = Array(entMaxPt(0), entMaxPt(1))
        corners(3) = Array(entMinPt(0), entMaxPt(1))
        
        ' Rotate each corner and update bounds
        Dim i As Long
        For i = 0 To 3
            Dim x As Double: x = corners(i)(0) - centerPt(0)
            Dim y As Double: y = corners(i)(1) - centerPt(1)
            
            Dim rx As Double: rx = x * cosA - y * sinA
            Dim ry As Double: ry = x * sinA + y * cosA
            
            If rx < minX Then minX = rx
            If rx > maxX Then maxX = rx
            If ry < minY Then minY = ry
            If ry > maxY Then maxY = ry
        Next i
        
        On Error GoTo 0
    Next ent
    
    minPt = Array(minX, minY, 0)
    maxPt = Array(maxX, maxY, 0)
End Sub

' Rotate all entities in the selection set
Private Sub RotateSelectedEntities(ss As AcadSelectionSet, centerPt() As Double, angle As Double)
    Dim ent As AcadEntity
    Dim errCount As Long: errCount = 0
    
    For Each ent In ss
        On Error Resume Next
        ent.Rotate centerPt, angle
        If Err.Number <> 0 Then
            errCount = errCount + 1
            Err.Clear
        End If
        On Error GoTo 0
    Next ent
    
    If errCount > 0 Then
        MsgBox "Aviso: " & errCount & " entidade(s) nao pudo ser rotacionada(s).", vbExclamation
    End If
End Sub
