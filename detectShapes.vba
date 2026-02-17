Option Explicit

' ------------------------------------------------------------------------------
' Module: detectShapes.vba
' Description: Performs Topological Data Analysis to detect outer boundaries
'              within a user selection and highlights them.
' Usage: Called from formPerfisul01
' ------------------------------------------------------------------------------

Public Sub DetectOuterShapes()
    ' 1. Hide the userform
    On Error Resume Next
    formPerfisul01.Hide
    On Error GoTo 0
    
    Dim doc As AcadDocument
    Set doc = ThisDrawing
    
    ' 2. Prompt for area selection
    ' We use a selection set to gather objects
    Dim sSet As AcadSelectionSet
    On Error Resume Next
        Set sSet = doc.SelectionSets.Item("TDA_SELECTION")
        If Err.Number = 0 Then
            sSet.Delete
        End If
        Err.Clear
    On Error GoTo 0
    
    Set sSet = doc.SelectionSets.Add("TDA_SELECTION")
    
    ' Prompt via command line
    doc.Utility.Prompt vbCrLf & "Select objects for Outer Shape Detection (Window Selection recommended)..." & vbCrLf
    
    ' Allow user to select objects
    ' Filter for curves that can form regions: Lines, Arcs, Circles, Polylines, Splines, Ellipses
    Dim gpCode(0) As Integer
    Dim dataValue(0) As Variant
    gpCode(0) = 0
    dataValue(0) = "LINE,ARC,CIRCLE,LWPOLYLINE,POLYLINE,SPLINE,ELLIPSE"
    
    On Error Resume Next
    sSet.SelectOnScreen gpCode, dataValue
    If sSet.Count = 0 Then
        MsgBox "Selection canceled or no valid objects selected.", vbExclamation
        If Not formPerfisul01 Is Nothing Then formPerfisul01.Show
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 3. Topological Data Analysis (Outer Shape Detection)
    ' Strategy:
    ' 1. Attempt to create Regions from the selection. This automatically finds closed loops.
    ' 2. Regions effectively represent the topology. Use boolean logic to find parent loops.
    '    A "Parent" or "Outer" loop is one that is NOT contained within any other region in the set.
    
    ' Collect objects for AddRegion
    Dim objs() As Object
    ReDim objs(sSet.Count - 1)
    Dim i As Integer
    For i = 0 To sSet.Count - 1
        Set objs(i) = sSet.Item(i)
    Next i
    
    ' Create Regions
    ' Note: AddRegion might fail if curves don't form closed loops.
    ' It returns a Variant containing an array of Region objects.
    Dim createdRegions As Variant
    On Error Resume Next
    createdRegions = doc.ModelSpace.AddRegion(objs)
    If Err.Number <> 0 Then
        MsgBox "Failed to create regions. Ensure selected objects form closed loops.", vbCritical
        If Not formPerfisul01 Is Nothing Then formPerfisul01.Show
        Exit Sub
    End If
    On Error GoTo 0
    
    If IsEmpty(createdRegions) Then
        MsgBox "No closed regions could be detected from selection.", vbExclamation
        If Not formPerfisul01 Is Nothing Then formPerfisul01.Show
        Exit Sub
    End If
    
    ' Store regions in an array for easier handling
    ' We need to cast the Variant array to something strictly typed or just iterate
    Dim regionList() As AcadRegion
    Dim rCount As Integer
    rCount = UBound(createdRegions) - LBound(createdRegions) + 1
    ReDim regionList(rCount - 1)
    
    Dim r As Variant
    Dim idx As Integer
    idx = 0
    For Each r In createdRegions
        Set regionList(idx) = r
        idx = idx + 1
    Next r
    
    ' Sort regionList by Area Descending (Bubble Sort)
    ' This optimization helps because a region can only be inside a LARGER region.
    Dim tempReg As AcadRegion
    Dim sorted As Boolean
    Dim j As Integer
    sorted = False
    While Not sorted
        sorted = True
        For i = 0 To UBound(regionList) - 1
            If regionList(i).Area < regionList(i + 1).Area Then
                Set tempReg = regionList(i)
                Set regionList(i) = regionList(i + 1)
                Set regionList(i + 1) = tempReg
                sorted = False
            End If
        Next i
    Wend
    
    ' Identify Outer Shapes
    Dim isInner As Boolean
    Dim k As Integer
    Dim testReg As AcadRegion
    Dim containerReg As AcadRegion
    Dim copyTest As AcadRegion
    Dim copyContainer As AcadRegion
    Dim intersectionReg As AcadRegion
    
    doc.Utility.Prompt "Analyzing " & rCount & " detected regions..." & vbCrLf
    
    For i = 0 To UBound(regionList)
        Set testReg = regionList(i)
        isInner = False
        
        ' Compare against other regions to see if 'testReg' is inside 'containerReg'
        For k = 0 To UBound(regionList)
            If i <> k Then
                Set containerReg = regionList(k)
                
                ' Only check if container is larger (or equal)
                ' Since we sorted descending, effectively we only check k < i?
                ' Actually, Equal areas are tricky (duplicate regions).
                ' If k < i, Area(k) >= Area(i).
                ' If Area(k) < Area(i), it cannot contain i.
                If containerReg.Area >= testReg.Area Then
                     
                    ' Check containment: Intersection(Test, Container) == Test
                    ' Use Copys to avoid altering originals
                    Set copyTest = testReg.Copy()
                    Set copyContainer = containerReg.Copy()
                    
                    On Error Resume Next
                    ' Boolean(acIntersection, object) modifies the object calling the method
                    copyTest.Boolean acIntersection, copyContainer
                    
                    If Err.Number = 0 Then
                         ' Check if area matches original area
                         If Abs(copyTest.Area - testReg.Area) < 0.0001 Then
                             ' It is inside!
                             isInner = True
                             
                             ' Clean up copies
                             copyTest.Delete
                             copyContainer.Delete
                             Exit For
                         End If
                    End If
                    On Error GoTo 0
                    
                    ' Clean up copies
                    If Not copyTest Is Nothing Then 
                        On Error Resume Next 
                        copyTest.Delete 
                    End If
                    If Not copyContainer Is Nothing Then 
                        On Error Resume Next 
                        copyContainer.Delete 
                    End If
                    On Error GoTo 0
                End If
            End If
        Next k
        
        If Not isInner Then
            ' It is a Parent / Outer Shape
            ' Highlight the upper-level/parent loops INSIDE selection.
            testReg.Color = acGreen
            On Error Resume Next
            testReg.Linetype = "Continuous"
            On Error GoTo 0
            testReg.Highlight True
        Else
            ' It is an inner loop / hole
            ' Delete the generated region to keep drawing clean?
            ' Or keep it? "Highlight parent" implies emphasizing them.
            ' Deleting non-parents makes it clearer.
            testReg.Delete
        End If
    Next i
    
    doc.Utility.Prompt "Analysis Complete. Outer shapes highlighted in RED." & vbCrLf
    MsgBox "Analysis Complete. Outer shapes highlighted.", vbInformation
    
    ' Show userform again
    If Not formPerfisul01 Is Nothing Then formPerfisul01.Show

End Sub
