Option Explicit

' Scales a user selection so that its largest single-line text height
' matches a user-supplied target height (in mm).
'
' Supported text types for height detection: AcDbText, AcDbAttribute,
' AcDbAttributeDefinition.  MText (AcDbMText) is intentionally skipped
' because its .Height property returns the bounding-box height, not the
' individual character height.

Public Sub ScaleToTextSize()
    Dim doc As AcadDocument
    Dim ss As AcadSelectionSet
    Dim ent As AcadEntity
    Dim i As Long
    Dim h As Double
    Dim maxTextHeight As Double
    Dim targetHeight As Double
    Dim scaleFactor As Double
    Dim bboxMin As Variant
    Dim bboxMax As Variant
    Dim totalMinPt(0 To 2) As Double
    Dim totalMaxPt(0 To 2) As Double
    Dim centerPt(0 To 2) As Double
    Dim firstEnt As Boolean
    Dim formWasVisible As Boolean
    Dim strTarget As String

    Set doc = ThisDrawing

    ' ---------- 0. Hide userform so it does not block the viewport ----------
    formWasVisible = False
    On Error Resume Next
    formWasVisible = formPerfisul01.Visible
    If formWasVisible Then formPerfisul01.Hide
    On Error GoTo 0

    ' ---------- Main loop: repeat until user cancels ----------
    Do
        ' ----- 1. Prompt user to area-select objects -----
        On Error Resume Next
        doc.SelectionSets("STS_SS").Delete
        On Error GoTo 0

        Set ss = doc.SelectionSets.Add("STS_SS")
        doc.Utility.Prompt vbCr & "Select objects to scale (empty selection to exit): "
        ss.SelectOnScreen

        If ss.Count = 0 Then
            ss.Delete
            Exit Do
        End If

        ' ----- 2. Find the largest text height in the selection -----
        maxTextHeight = 0
        For i = 0 To ss.Count - 1
            Set ent = ss.Item(i)
            h = 0
            Select Case ent.ObjectName
                Case "AcDbText", "AcDbAttribute", "AcDbAttributeDefinition"
                    On Error Resume Next
                    h = ent.Height
                    On Error GoTo 0
            End Select
            If h > maxTextHeight Then maxTextHeight = h
        Next i

        If maxTextHeight = 0 Then
            MsgBox "No text objects (TEXT / ATTDEF / ATTRIB) found in selection." & vbCr & _
                   "Please select again.", vbExclamation
            ss.Delete

        Else
            ' ----- 3. Prompt for target text height (InputBox) -----
            strTarget = InputBox( _
                "Largest text height found:  " & Format(maxTextHeight, "0.####") & " mm" & vbCr & vbCr & _
                "Enter the target text height (mm):", _
                "Scale to Text Size", _
                Format(maxTextHeight, "0.####"))

            If strTarget = "" Then
                ' User cancelled the dialog
                ss.Delete
                Exit Do
            End If

            On Error Resume Next
            targetHeight = CDbl(strTarget)
            If Err.Number <> 0 Or targetHeight <= 0 Then
                On Error GoTo 0
                MsgBox "Invalid value. Please enter a number greater than zero.", vbExclamation
                ss.Delete

            Else
                On Error GoTo 0

                ' ----- 4. Compute the bounding-box centre of the full selection -----
                firstEnt = True
                For i = 0 To ss.Count - 1
                    Set ent = ss.Item(i)
                    On Error Resume Next
                    ent.GetBoundingBox bboxMin, bboxMax
                    If Err.Number = 0 Then
                        If firstEnt Then
                            totalMinPt(0) = bboxMin(0): totalMinPt(1) = bboxMin(1): totalMinPt(2) = 0
                            totalMaxPt(0) = bboxMax(0): totalMaxPt(1) = bboxMax(1): totalMaxPt(2) = 0
                            firstEnt = False
                        Else
                            If bboxMin(0) < totalMinPt(0) Then totalMinPt(0) = bboxMin(0)
                            If bboxMin(1) < totalMinPt(1) Then totalMinPt(1) = bboxMin(1)
                            If bboxMax(0) > totalMaxPt(0) Then totalMaxPt(0) = bboxMax(0)
                            If bboxMax(1) > totalMaxPt(1) Then totalMaxPt(1) = bboxMax(1)
                        End If
                    End If
                    On Error GoTo 0
                Next i

                If firstEnt Then
                    MsgBox "Could not compute selection bounding box. Please select again.", vbExclamation
                    ss.Delete

                Else
                    centerPt(0) = (totalMinPt(0) + totalMaxPt(0)) / 2
                    centerPt(1) = (totalMinPt(1) + totalMaxPt(1)) / 2
                    centerPt(2) = 0

                    ' ----- 5. Scale all selected entities about the selection centre -----
                    scaleFactor = targetHeight / maxTextHeight

                    doc.StartUndoMark

                    For i = 0 To ss.Count - 1
                        Set ent = ss.Item(i)
                        On Error Resume Next
                        ent.ScaleEntity centerPt, scaleFactor
                        On Error GoTo 0
                    Next i

                    doc.EndUndoMark

                    doc.Utility.Prompt vbCr & _
                        "Done. " & ss.Count & " object(s) scaled by " & Format(scaleFactor, "0.####") & _
                        "  (text height: " & Format(maxTextHeight, "0.##") & " mm  →  " & _
                        Format(targetHeight, "0.##") & " mm)." & vbCr

                    ss.Delete
                End If
            End If
        End If
    Loop

Cleanup:
    ' Restore userform visibility
    If formWasVisible Then
        On Error Resume Next
        formPerfisul01.Show vbModeless
        On Error GoTo 0
    End If
End Sub
