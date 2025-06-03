
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
                ' Add more cases as needed for other entity types
            End Select
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
