
' Helper to copy text to clipboard using Windows API (no MSForms required)
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)

Private Const CF_TEXT As Long = 1
Private Const GMEM_MOVEABLE As Long = &H2


Public Sub SumTotalPerimeter()
    Dim doc As AcadDocument
    Dim ent As AcadEntity
    Dim totalPerimeter As Double
    Dim strResult As String
    Dim layerName As String
    Dim i As Long

    Set doc = ThisDrawing

    ' Get layer name from combobox cbCamada02
    On Error Resume Next
    layerName = formPerfisul01.cbCamada02.Value
    On Error GoTo 0
    If layerName = "" Then
        MsgBox "No layer selected in cbCamada02. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    totalPerimeter = 0

    ' Loop through all entities in ModelSpace
    For i = 0 To doc.ModelSpace.Count - 1
        Set ent = doc.ModelSpace.Item(i)
        If UCase(ent.Layer) = UCase(layerName) Then
            On Error Resume Next
            Select Case ent.ObjectName
                Case "AcDbLine"
                    totalPerimeter = totalPerimeter + ent.Length
                Case "AcDbPolyline", "AcDb2dPolyline", "AcDb3dPolyline"
                    totalPerimeter = totalPerimeter + ent.Length
                Case "AcDbCircle"
                    totalPerimeter = totalPerimeter + ent.Circumference
                Case "AcDbArc"
                    totalPerimeter = totalPerimeter + ent.ArcLength
                Case "AcDbEllipse"
                    totalPerimeter = totalPerimeter + ent.ArcLength
                Case "AcDbSpline"
                    ' Calculate spline length by sampling points along the curve
                    Dim spline As AcadSpline
                    Dim splineLen As Double
                    Dim pt1 As Variant, pt2 As Variant
                    Dim numSamples As Integer
                    Dim j As Integer
                    Dim t As Double, tStart As Double, tEnd As Double
                    
                    Set spline = ent
                    splineLen = 0
                    numSamples = 100
                    
                    ' Get start and end fit points to estimate parameter range
                    Dim fitPts As Variant
                    Dim ctrlPts As Variant
                    ctrlPts = spline.ControlPoints
                    
                    ' Approximate length by summing distances between sampled points
                    tStart = 0
                    tEnd = 1
                    
                    On Error Resume Next
                    pt1 = spline.GetPointAtParam(tStart)
                    If Err.Number <> 0 Then
                        ' Try using control points to approximate
                        Err.Clear
                        Dim k As Integer
                        For k = 0 To (UBound(ctrlPts) - 3) Step 3
                            splineLen = splineLen + Sqr((ctrlPts(k + 3) - ctrlPts(k)) ^ 2 + _
                                                        (ctrlPts(k + 4) - ctrlPts(k + 1)) ^ 2 + _
                                                        (ctrlPts(k + 5) - ctrlPts(k + 2)) ^ 2)
                        Next k
                    Else
                        ' Sample points along parameter range
                        For j = 1 To numSamples
                            t = tStart + (tEnd - tStart) * (j / numSamples)
                            pt2 = spline.GetPointAtParam(t)
                            If Err.Number = 0 Then
                                splineLen = splineLen + Sqr((pt2(0) - pt1(0)) ^ 2 + _
                                                            (pt2(1) - pt1(1)) ^ 2 + _
                                                            (pt2(2) - pt1(2)) ^ 2)
                                pt1 = pt2
                            End If
                        Next j
                    End If
                    On Error Resume Next
                    
                    totalPerimeter = totalPerimeter + splineLen
                ' Add more cases as needed for other entity types
            End Select
            On Error GoTo 0
        End If
    Next i

    strResult = "Total Perimeter on layer '" & layerName & "': " & Format(totalPerimeter, "0.00")
    MsgBox strResult, vbInformation, "Sum of Perimeters"

    ' Copy to clipboard
    CopyTextToClipboard Format(totalPerimeter, "0.00")
End Sub


Public Sub CopyTextToClipboard(ByVal sText As String)
    Dim hGlobalMemory As LongPtr
    Dim lpGlobalMemory As LongPtr
    Dim hWnd As LongPtr
    Dim sData As String
    Dim lSize As Long

    sData = sText & vbNullChar
    lSize = LenB(sData)

    hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE, lSize)
    If hGlobalMemory Then
        lpGlobalMemory = GlobalLock(hGlobalMemory)
        If lpGlobalMemory Then
            CopyMemory lpGlobalMemory, StrPtr(sData), lSize
            GlobalUnlock hGlobalMemory
            If OpenClipboard(0&) Then
                EmptyClipboard
                SetClipboardData CF_TEXT, hGlobalMemory
                CloseClipboard
            End If
        End If
    End If
End Sub
