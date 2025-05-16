Option Explicit

' MText_To_Text: Convert MText objects on a specified layer to Text objects
' This preserves text content, location, height, rotation and layer
' Use the layer selection from formPerfisul01

Public Sub MText_To_Text(Optional ByVal layerNameParam As String = "")
    ' Variable declarations
    Dim acadDoc As Object ' AcadDocument
    Dim modelSpace As Object ' AcadModelSpace
    Dim entity As Object ' AcadEntity
    Dim mtextObj As Object ' AcadMText
    Dim layerName As String
    Dim mtextObjectsOnLayer As Collection
    Dim i As Long
    Dim mTextCount As Long
    Dim conversionCount As Long
    
    ' --- Form Access and Input Validation ---
    If layerNameParam = "" Then
        On Error Resume Next
        layerName = formPerfisul01.cbCamada01.Value
        If Err.Number <> 0 Then
            MsgBox "Erro ao acessar o controle de camada no formulário.", vbCritical, "Erro de Controle"
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
        If layerName = "-- Selecione uma camada --" Or layerName = "" Then
            MsgBox "Por favor, selecione uma camada válida.", vbExclamation, "Seleção de Camada"
            Exit Sub
        End If
    Else
        layerName = layerNameParam
    End If
    
    On Error GoTo ErrorHandler
    
    ' Get the current document and model space
    Set acadDoc = ThisDrawing
    Set modelSpace = acadDoc.ModelSpace
    
    ' --- Search for MText objects in the specified layer ---
    Set mtextObjectsOnLayer = New Collection
    For Each entity In modelSpace
        If entity.Layer = layerName Then
            If entity.ObjectName = "AcDbMText" Then
                mtextObjectsOnLayer.Add entity
            End If
        End If
    Next entity
    
    mTextCount = mtextObjectsOnLayer.Count
    If mTextCount = 0 Then
      '   MsgBox "Nenhum objeto MText encontrado na camada '" & layerName & "'.", vbInformation
        GoTo Cleanup
    End If
    
    ' Show progress message
   '  MsgBox "Encontrados " & mTextCount & " objetos MText na camada '" & layerName & "'." & vbCrLf & _
   '         "Tentando convertê-los para objetos de texto...", vbInformation
      ' --- Try to explode MText objects first ---
    ' This is a more reliable way to convert MText to Text in BricsCAD
    ' Process each MText object individually for greater reliability
    conversionCount = 0
    
    For Each mtextObj In mtextObjectsOnLayer        ' Build a simple explode command for each object
        Dim explodeCmd As String
        On Error Resume Next
        
        ' Direct selection and explode of just this one MText object
        explodeCmd = "_SELECT (handent """ & mtextObj.Handle & """) " & vbCr & "_EXPLODE " & vbCr
        
        ' Execute the commands
        acadDoc.SendCommand explodeCmd
        
        If Err.Number = 0 Then
            conversionCount = conversionCount + 1
        Else
            Debug.Print "Error exploding MText: " & mtextObj.Handle & " - " & Err.Description
            Err.Clear
        End If
    Next mtextObj
    
    ' Update display after all operations
    acadDoc.Regen acAllViewports
          ' Pause to let BricsCAD process the commands
        Wait 1
        
        ' Check if explosion approach worked
        If conversionCount > 0 Then
            ' MsgBox conversionCount & " de " & mTextCount & " objetos MText foram explodidos na camada '" & _
                   layerName & "'.", vbInformation
              ' Check if we need to continue with alternate approach
            If conversionCount >= mTextCount Then
                GoTo Cleanup ' All objects were processed successfully
            End If
        End If
        
        ' If we get here, some objects remain - fall back to manual conversion
    
    ' --- Fall back to manual conversion if explosion didn't work ---
    conversionCount = 0
    
    ' Re-collect MText objects as some might have been exploded
    Set mtextObjectsOnLayer = New Collection
    For Each entity In modelSpace
        If entity.Layer = layerName Then
            If entity.ObjectName = "AcDbMText" Then
                mtextObjectsOnLayer.Add entity
            End If
        End If
    Next entity
    
    mTextCount = mtextObjectsOnLayer.Count
    If mTextCount = 0 Then
      '   MsgBox "Todos os objetos MText foram explodidos com sucesso.", vbInformation
        GoTo Cleanup
    End If
    
    ' Try alternative method: create text and delete originals
    For i = mTextCount To 1 Step -1 ' Process backwards for collection integrity
        Set mtextObj = mtextObjectsOnLayer(i)
        
        On Error Resume Next
        ' Get original MText properties
        Dim mtextContent As String
        Dim mtextInsertionPoint As Variant
        Dim mtextHeight As Double
        Dim mtextRotation As Double
        Dim mtextColor As Integer
        Dim mtextLayer As String
        Dim mtextStyle As String
        
        mtextContent = CleanMTextString(mtextObj.Text) ' Use the helper function to clean formatting
        mtextInsertionPoint = mtextObj.InsertionPoint
        mtextHeight = mtextObj.Height
        mtextRotation = 0 ' Default in case we can't get rotation
        
        On Error Resume Next
        mtextRotation = mtextObj.Rotation
        If Err.Number <> 0 Then
            Err.Clear
        End If
        
        mtextColor = mtextObj.Color
        mtextLayer = mtextObj.Layer
        mtextStyle = mtextObj.StyleName
        
        ' Create new Text object with the same properties
        Dim newTextObj As Object
        Set newTextObj = modelSpace.AddText(mtextContent, mtextInsertionPoint, mtextHeight)
        newTextObj.Rotation = mtextRotation
        newTextObj.Layer = mtextLayer
        newTextObj.StyleName = mtextStyle
        newTextObj.Color = mtextColor
        newTextObj.Update        ' Remove the original MText object using a targeted selection
        Dim eraseCmd As String
        eraseCmd = "_SELECT (handent """ & mtextObj.Handle & """) " & vbCr & "_ERASE " & vbCr
        
        On Error Resume Next
        acadDoc.SendCommand eraseCmd
        
        ' Short pause to allow command to complete
        Wait 0.2
        
        ' Verify the object was deleted
        Dim deleted As Boolean
        deleted = False
        Dim testHandle As String
        testHandle = mtextObj.Handle
        If Err.Number <> 0 Then
            ' Error accessing handle means the object was deleted
            deleted = True
            conversionCount = conversionCount + 1
            Err.Clear
        End If
        
        If Not deleted Then
            ' Try an alternative direct approach if the first method failed - use API method
            On Error Resume Next
            mtextObj.Delete  ' Try the API method
            Wait 0.2
            Wait 0.2
            conversionCount = conversionCount + 1
        End If
        
        On Error GoTo ErrorHandler ' Restore main error handler
    Next i
    
    ' Update the display and inform the user
    acadDoc.Regen acAllViewports
    MsgBox conversionCount & " de " & mTextCount & " objetos MText convertidos para Texto na camada '" & _
           layerName & "'.", vbInformation
    
Cleanup:
    Set acadDoc = Nothing
    Set modelSpace = Nothing
    Set mtextObjectsOnLayer = Nothing
    Set entity = Nothing
    Set mtextObj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Ocorreu um erro: " & Err.Description & vbCrLf & _
           "Número do erro: " & Err.Number & vbCrLf & _
           "Fonte do erro: " & Err.Source, vbCritical, "Erro de Script"
    Resume Cleanup
End Sub

' --- Helper function to clean MText formatting ---
' Used to strip formatting from MText content for cleaner conversion to Text
Private Function CleanMTextString(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    
    ' Remove {...} formatting codes
    re.Pattern = "\\{[^}]*\\}"
    s = re.Replace(s, "")
    
    ' Replace CR, LF, Tab, No-break space with a single space
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr(160), " ")
    
    ' Normalize multiple spaces to a single space
    re.Pattern = "\\s+" ' Match one or more whitespace characters
    s = re.Replace(s, " ")
    
    s = Trim(s) ' Trim leading/trailing spaces
    CleanMTextString = s
End Function

' Helper subroutine to add a small delay (in seconds)
Private Sub Wait(seconds As Double)
    Dim startTime As Double
    startTime = Timer
    Do While Timer < startTime + seconds
        DoEvents ' Allow the system to process other events
    Loop
End Sub

' Instructions to run:
' 1. Open BricsCAD.
' 2. Press ALT+F11 to open the VBA IDE (or type VBAIDE in the command line).
' 3. In the VBA IDE, go to File > Import File... and select this .vba file.
'    Alternatively, go to Insert > Module and paste this code into the new module.
' 4. Close the VBA IDE.
' 5. In BricsCAD, type VBARUN in the command line.
' 6. Select "MText_To_Text" from the list and click "Run".
' 7. Ensure that formPerfisul01 is loaded and a layer is selected in cbCamada01.