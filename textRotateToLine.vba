Option Explicit

Public Sub RotateTextToLineAngle()
    On Error GoTo ErrorHandler
    
    Dim acadApp As AcadApplication
    Dim acadDoc As AcadDocument
    Dim objLine As Object ' AcadLine
    Dim objText As Object ' AcadText or AcadMText
    Dim pickPt As Variant
    Dim lineAngle As Double
    Dim midPt As Variant
    
    Set acadApp = ThisDrawing.Application
    Set acadDoc = ThisDrawing
    
    Do While True
        ' Select the line
        Err.Clear
        On Error Resume Next
        acadDoc.Utility.GetEntity objLine, pickPt, vbCrLf & "Select a line (<Esc> to stop): "
        If Err.Number <> 0 Then Exit Do ' Esc cancels the loop
        On Error GoTo ErrorHandler
        
        If objLine.ObjectName <> "AcDbLine" Then
            MsgBox "Please select a line entity.", vbExclamation
            GoTo ContinueLoop
        End If
        midPt = MidPoint(objLine.StartPoint, objLine.EndPoint)
        
        ' Select the text (Text or MText)
        Err.Clear
        On Error Resume Next
        acadDoc.Utility.GetEntity objText, pickPt, vbCrLf & "Select a text (TEXT/MTEXT) (<Esc> to stop): "
        If Err.Number <> 0 Then Exit Do ' Esc cancels the loop
        On Error GoTo ErrorHandler
        
        If Not (objText.ObjectName = "AcDbText" Or objText.ObjectName = "AcDbMText") Then
            MsgBox "Please select a TEXT or MTEXT entity.", vbExclamation
            GoTo ContinueLoop
        End If
        
        lineAngle = objLine.Angle ' radians
        objText.Rotation = lineAngle

        ' Place text at line midpoint and align bottom-center
        If objText.ObjectName = "AcDbText" Then
            objText.Alignment = acAlignmentBottomCenter
            objText.TextAlignmentPoint = midPt
            objText.InsertionPoint = midPt
        ElseIf objText.ObjectName = "AcDbMText" Then
            objText.AttachmentPoint = acAttachmentPointBottomCenter
            objText.InsertionPoint = midPt
        End If
        
        acadDoc.Regen acAllViewports
ContinueLoop:
    Loop
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function MidPoint(p1 As Variant, p2 As Variant) As Variant
    Dim m(0 To 2) As Double
    m(0) = (p1(0) + p2(0)) / 2
    m(1) = (p1(1) + p2(1)) / 2
    m(2) = (p1(2) + p2(2)) / 2
    MidPoint = m
End Function
