Option Explicit

' ==============================================================================
' Module  : celloFingerboard.vba
' Description:
'   Draws a cello fingerboard diagram in BricsCAD ModelSpace.
'
'   The diagram contains:
'     • Four lines representing the four cello strings (C2, G2, D3, A3).
'     • Small circles marking every finger position using Equal Temperament
'       tuning distances from the nut.
'
'   Equal Temperament formula (distance from nut to semitone n):
'     d = L × ( 1 − 2^(−n/12) )
'     where L = scale length (mm), n = semitone offset above the open string.
'
'   Color modes (USE_STRING_COLORS):
'     True  → each string has its own fixed color.
'     False → each note name (C, C#, D … B) carries its own color class.
'
'   Sharp / flat control (DRAW_SHARPS):
'     True  → C#, D#, F#, G#, A# positions are drawn on CELLO_Sharps layer.
'     False → only natural notes (C, D, E, F, G, A, B) are drawn.
'
'   Layers created automatically:
'     CELLO_Strings  – the four string lines.
'     CELLO_Naturals – natural-note circles.
'     CELLO_Sharps   – sharp / flat circles (created only when DRAW_SHARPS is True).
'
'   Entry point:  DrawCelloFingerboard
' ==============================================================================


' ==============================================================================
' SECTION 1 – INSTRUMENT DIMENSIONS
' Adjust these values to match any cello size or alternate string instrument.
' ==============================================================================

Private Const SCALE_LENGTH       As Double = 690#   ' mm  Full vibrating string length (nut to bridge)
Private Const FINGERBOARD_LEN    As Double = 250#   ' mm  Visible fingerboard length to render
Private Const NUT_STRING_SPAN    As Double = 33#    ' mm  Outer-edge span of all 4 strings measured at the nut
Private Const BRIDGE_STRING_SPAN As Double = 90#    ' mm  Outer-edge span of all 4 strings measured at the bridge


' ==============================================================================
' SECTION 2 – DRAWING LAYOUT
' Controls the position and size of drawn geometry.
' ==============================================================================

Private Const ORIGIN_X           As Double = 0#     ' mm  X reference for the leftmost string at the nut
Private Const ORIGIN_Y           As Double = 0#     ' mm  Y reference for the nut line (top of diagram; bridge is below)
Private Const NOTE_RADIUS        As Double = 2.5    ' mm  Radius of each finger-position circle
Private Const NUM_SEMITONES      As Integer = 24    ' #   Semitone positions drawn per string (24 = two octaves)
Private Const STRING_WIDTH_SCALE As Double = 1#     ' ×   Multiplier applied to realistic string diameters (1.0 = true to life)
Private Const LEGEND_OFFSET_X    As Double = 50#    ' mm  Gap from the fingerboard right edge to the legend circle centre
Private Const LEGEND_ROW_SPACING As Double = 8#     ' mm  Vertical distance between legend rows (centre to centre)
Private Const LEGEND_TEXT_HEIGHT As Double = 3.5    ' mm  Note label text height


' ==============================================================================
' SECTION 3 – DISPLAY OPTIONS
' Toggle these flags to control what is drawn and how colors are applied.
' ==============================================================================

Private Const DRAW_SHARPS        As Boolean = False  ' Bool  True  = include sharp/flat (black-key) note circles
                                                    '       False = natural notes only
Private Const USE_STRING_COLORS  As Boolean = False  ' Bool  True  = color circles by string identity
                                                    '       False = color circles by note name (chromatic class)


' ==============================================================================
' SECTION 4 – LAYER NAMES
' Geometry is distributed across separate layers for easy visibility control.
' ==============================================================================

Private Const LAYER_STRINGS      As String = "CELLO_Strings"   ' String lines
Private Const LAYER_NATURALS     As String = "CELLO_Naturals"  ' Natural-note circles
Private Const LAYER_SHARPS       As String = "CELLO_Sharps"    ' Sharp/flat note circles
Private Const LAYER_LEGEND       As String = "CELLO_Legend"    ' Colour legend entities
Private Const LEGEND_TEXT_STYLE  As String = "CELLO_Arial"     ' Arial text style used by the legend


' ==============================================================================
' PUBLIC ENTRY POINT
' ==============================================================================

' DrawCelloFingerboard
' --------------------
' Main macro entry point. Call this from the BricsCAD Macro dialog or a toolbar button.
'
' Workflow:
'   1. Ensure all drawing layers exist.
'   2. Draw the four cello string lines (C2, G2, D3, A3).
'   3. Plot finger-position circles for every requested semitone on each string.
'   4. Regen the viewports.
Public Sub DrawCelloFingerboard()

    Dim doc As AcadDocument
    Set doc = ThisDrawing

    ' --- Step 1: Ensure required layers exist ---
    Call EnsureLayer(doc, LAYER_STRINGS, acWhite)
    Call EnsureLayer(doc, LAYER_NATURALS, acWhite)
    If DRAW_SHARPS Then
        Call EnsureLayer(doc, LAYER_SHARPS, acWhite)
    End If
    Call EnsureLayer(doc, LAYER_LEGEND, acWhite)
    Call EnsureTextStyle(doc, LEGEND_TEXT_STYLE, "arial.ttf")

    ' --- Step 2: Draw the four string lines (indices 0=C, 1=G, 2=D, 3=A) ---
    Dim s As Integer
    For s = 0 To 3
        Call DrawStringLine(doc, s)
    Next s

    ' --- Step 3: Draw finger-position circles on each string ---
    For s = 0 To 3
        Call DrawNoteCircles(doc, s)
    Next s

    ' --- Step 4: Draw colour / note legend ---
    Call DrawColorLegend(doc)

    ' --- Step 5: Refresh all viewports ---
    doc.Regen acAllViewports

    ' Summary message
    Dim msg As String
    msg = "Cello fingerboard diagram drawn successfully." & vbCrLf & vbCrLf
    msg = msg & "  Scale length      : " & SCALE_LENGTH & " mm" & vbCrLf
    msg = msg & "  Fingerboard length: " & FINGERBOARD_LEN & " mm" & vbCrLf
    msg = msg & "  Semitones drawn   : " & NUM_SEMITONES & " per string" & vbCrLf
    msg = msg & "  Sharps drawn      : " & IIf(DRAW_SHARPS, "Yes", "No") & vbCrLf
    msg = msg & "  Color mode        : " & IIf(USE_STRING_COLORS, "By string", "By note name")
    MsgBox msg, vbInformation, "DrawCelloFingerboard"

End Sub


' ==============================================================================
' DRAWING – STRING LINES
' ==============================================================================

' StringBaseWidthMM
' -----------------
' Returns the realistic physical diameter (mm) of each cello string.
' Values are typical wound-steel/gut-core gauges for a full-size cello.
' Multiply by STRING_WIDTH_SCALE for fine-tuning.
'
'   stringIndex : 0 = C2  1 = G2  2 = D3  3 = A3
Private Function StringBaseWidthMM(stringIndex As Integer) As Double
    Select Case stringIndex
        Case 0: StringBaseWidthMM = 1.1    ' C2 – thickest wound string
        Case 1: StringBaseWidthMM = 0.9    ' G2
        Case 2: StringBaseWidthMM = 0.7    ' D3
        Case 3: StringBaseWidthMM = 0.55   ' A3 – thinnest string
        Case Else: StringBaseWidthMM = 0.5
    End Select
End Function


' NearestLineweight
' -----------------
' Snaps a desired lineweight in mm to the nearest value supported by
' AutoCAD / BricsCAD (which only accept a fixed set of integer weights
' expressed in hundredths of a millimetre).
'
'   widthMM : desired lineweight in millimetres
Private Function NearestLineweight(widthMM As Double) As Long

    ' All valid ACI lineweight values in 1/100 mm units
    Dim lw(23) As Long
    lw(0)  = 0:   lw(1)  = 5:   lw(2)  = 9:   lw(3)  = 13
    lw(4)  = 15:  lw(5)  = 18:  lw(6)  = 20:  lw(7)  = 25
    lw(8)  = 30:  lw(9)  = 35:  lw(10) = 40:  lw(11) = 50
    lw(12) = 53:  lw(13) = 60:  lw(14) = 70:  lw(15) = 80
    lw(16) = 90:  lw(17) = 100: lw(18) = 106: lw(19) = 120
    lw(20) = 140: lw(21) = 158: lw(22) = 200: lw(23) = 211

    Dim target As Long
    target = CLng(widthMM * 100#)

    Dim best As Long
    Dim bestDiff As Long
    best = lw(0)
    bestDiff = Abs(target - lw(0))

    Dim i As Integer
    For i = 1 To 23
        Dim d As Long
        d = Abs(target - lw(i))
        If d < bestDiff Then
            bestDiff = d
            best = lw(i)
        End If
    Next i

    NearestLineweight = best
End Function


' DrawStringLine
' --------------
' Draws one of the four cello strings as a straight line from the nut to the
' end of the visible fingerboard. The string X-coordinate is interpolated
' between the nut span and the bridge span to reflect the natural taper.
'
'   doc         : active AcadDocument object
'   stringIndex : 0 = C2 (lowest/widest)  1 = G2  2 = D3  3 = A3 (highest)
Private Sub DrawStringLine(doc As AcadDocument, stringIndex As Integer)

    ' Resolve the open-string chromatic root so the line color matches its open note
    Dim openChromatic As Integer
    Select Case stringIndex
        Case 0: openChromatic = 0   ' C
        Case 1: openChromatic = 7   ' G
        Case 2: openChromatic = 2   ' D
        Case 3: openChromatic = 9   ' A
    End Select

    ' Start point at the nut (Y = ORIGIN_Y, top of diagram)
    Dim startPt(2) As Double
    startPt(0) = StringXAtY(stringIndex, 0#)
    startPt(1) = ORIGIN_Y
    startPt(2) = 0#

    ' End point at the far end of the visible fingerboard (below the nut)
    Dim endPt(2) As Double
    endPt(0) = StringXAtY(stringIndex, FINGERBOARD_LEN)
    endPt(1) = ORIGIN_Y - FINGERBOARD_LEN
    endPt(2) = 0#

    Dim lineObj As AcadLine
    Set lineObj = doc.ModelSpace.AddLine(startPt, endPt)
    lineObj.Layer = LAYER_STRINGS
    ' Use the same color dispatcher as note circles so lines always match
    lineObj.Color = CircleColor(stringIndex, openChromatic)
    ' Apply realistic string diameter scaled by STRING_WIDTH_SCALE
    lineObj.Lineweight = NearestLineweight(StringBaseWidthMM(stringIndex) * STRING_WIDTH_SCALE)

End Sub


' ==============================================================================
' DRAWING – FINGER-POSITION CIRCLES
' ==============================================================================

' DrawNoteCircles
' ---------------
' Iterates through all requested semitone positions on a given string and places
' a circle at each one. The Y-position of each circle is calculated with the
' Equal Temperament formula. Sharp notes are drawn only when DRAW_SHARPS = True.
' Positions beyond FINGERBOARD_LEN are silently skipped.
'
'   doc         : active AcadDocument object
'   stringIndex : 0 = C2  1 = G2  2 = D3  3 = A3
Private Sub DrawNoteCircles(doc As AcadDocument, stringIndex As Integer)

    ' Open-string chromatic root (0=C … 11=B) and MIDI base note for octave labels
    Dim openChromatic As Integer    ' position within a chromatic octave (mod 12)
    Dim midiBase      As Integer    ' MIDI note number of the open string

    Select Case stringIndex
        Case 0: openChromatic = 0:  midiBase = 36  ' C2  (MIDI 36)
        Case 1: openChromatic = 7:  midiBase = 43  ' G2  (MIDI 43)
        Case 2: openChromatic = 2:  midiBase = 50  ' D3  (MIDI 50)
        Case 3: openChromatic = 9:  midiBase = 57  ' A3  (MIDI 57)
    End Select

    Dim n As Integer
    For n = 1 To NUM_SEMITONES      ' n=0 is the open string (nut); start at 1

        ' Absolute chromatic index for this note (may exceed 11; use mod 12 for note name)
        Dim absChromatic As Integer
        absChromatic = openChromatic + n

        ' Determine whether this semitone is a sharp / flat position
        Dim sharp As Boolean
        sharp = IsSharpNote(absChromatic)

        ' Conditionally skip sharp/flat positions based on the DRAW_SHARPS flag
        If Not (sharp And Not DRAW_SHARPS) Then

            ' --- Equal Temperament: distance from nut to this semitone ---
            Dim distFromNut As Double
            distFromNut = NoteDistanceFromNut(SCALE_LENGTH, n)

            ' Skip positions that fall beyond the drawable fingerboard length
            If distFromNut <= FINGERBOARD_LEN Then

                ' Compute the XY centre of the circle
                ' Y decreases downward from the nut (mirrored layout)
                Dim cx As Double
                Dim cy As Double
                cx = StringXAtY(stringIndex, distFromNut)
                cy = ORIGIN_Y - distFromNut

                ' Choose layer (natural vs. sharp) and color
                Dim lyr As String
                lyr = IIf(sharp, LAYER_SHARPS, LAYER_NATURALS)

                Dim col As Long
                col = CircleColor(stringIndex, absChromatic)

                ' Draw the finger-position circle
                Call DrawNoteCircle(doc, cx, cy, col, lyr)

            End If
        End If

    Next n

End Sub


' DrawNoteCircle
' --------------
' Places a single circle in ModelSpace at the given centre coordinates.
'
'   doc     : active AcadDocument object
'   x, y    : centre of the circle in mm
'   col     : AutoCAD Color Index (ACI) to apply to the entity
'   lyrName : target layer name
Private Sub DrawNoteCircle(doc As AcadDocument, _
                           x As Double, y As Double, _
                           col As Long, lyrName As String)
    Dim center(2) As Double
    center(0) = x
    center(1) = y
    center(2) = 0#

    Dim circObj As AcadCircle
    Set circObj = doc.ModelSpace.AddCircle(center, NOTE_RADIUS)
    circObj.Layer = lyrName
    circObj.Color = col

    ' Fill the circle with a solid hatch in the same color
    Dim hatchObj As AcadHatch
    Set hatchObj = doc.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True)
    hatchObj.Layer = lyrName
    hatchObj.Color = col
    Dim boundary(0) As AcadEntity
    Set boundary(0) = circObj
    hatchObj.AppendOuterLoop boundary
    hatchObj.Evaluate

End Sub


' ==============================================================================
' GEOMETRY HELPERS
' ==============================================================================

' StringXAtY
' ----------
' Returns the X coordinate of a given string at a distance y from the nut.
' The fingerboard tapers from NUT_STRING_SPAN at the nut to BRIDGE_STRING_SPAN
' at the bridge. Both spans are centred around the same X midpoint, so the
' taper is symmetrical. Linear interpolation is used along the full SCALE_LENGTH.
'
'   stringIndex : 0 = C (leftmost)  1 = G  2 = D  3 = A (rightmost)
'   y           : distance in mm from the nut
Private Function StringXAtY(stringIndex As Integer, y As Double) As Double

    ' Horizontal centre reference (ensures symmetric taper)
    Dim nutCenter As Double
    nutCenter = ORIGIN_X + NUT_STRING_SPAN / 2#

    ' X of this string at the nut end and at the bridge end
    Dim xNut    As Double
    Dim xBridge As Double
    xNut    = nutCenter - NUT_STRING_SPAN    / 2# + CDbl(stringIndex) * (NUT_STRING_SPAN    / 3#)
    xBridge = nutCenter - BRIDGE_STRING_SPAN / 2# + CDbl(stringIndex) * (BRIDGE_STRING_SPAN / 3#)

    ' Linear interpolation: t=0 at nut, t=1 at bridge
    Dim t As Double
    t = y / SCALE_LENGTH

    StringXAtY = xNut + (xBridge - xNut) * t

End Function


' NoteDistanceFromNut
' -------------------
' Returns the distance in mm from the nut to the given semitone position,
' using the standard Equal Temperament formula:
'
'   d = L × ( 1 − 2^(−n/12) )
'
' where L is the scale length in mm and n is the semitone offset from the
' open string. Returns 0 for n ≤ 0 (open string / nut position).
'
'   scaleLength : vibrating string length in mm (e.g. 690 for a full-size cello)
'   semitone    : integer offset above the open string (1 = first semitone)
Private Function NoteDistanceFromNut(scaleLength As Double, semitone As Integer) As Double

    If semitone <= 0 Then
        NoteDistanceFromNut = 0#
        Exit Function
    End If

    ' Core Equal Temperament formula
    NoteDistanceFromNut = scaleLength * (1# - (1# / (2# ^ (CDbl(semitone) / 12#))))

End Function


' IsSharpNote
' -----------
' Returns True when the given absolute chromatic index resolves to a sharp
' or flat note within an octave.
'
' Sharp positions within a chromatic octave (mod 12):
'   1 = C#/Db   3 = D#/Eb   6 = F#/Gb   8 = G#/Ab   10 = A#/Bb
'
'   chromaticIndex : any integer; only its mod-12 value is examined
Private Function IsSharpNote(chromaticIndex As Integer) As Boolean

    Select Case ((chromaticIndex Mod 12) + 12) Mod 12
        Case 1, 3, 6, 8, 10
            IsSharpNote = True
        Case Else
            IsSharpNote = False
    End Select

End Function


' NoteName
' --------
' Returns the note name string (e.g. "C", "C#", "D") for any absolute
' chromatic index, using sharp spelling for accidentals.
'
'   absoluteChromatic : any integer; only its mod-12 value is used
Private Function NoteName(absoluteChromatic As Integer) As String

    Dim names(11) As String
    names(0) = "C":  names(1)  = "C#": names(2)  = "D":  names(3)  = "D#"
    names(4) = "E":  names(5)  = "F":  names(6)  = "F#": names(7)  = "G"
    names(8) = "G#": names(9)  = "A":  names(10) = "A#": names(11) = "B"

    NoteName = names(((absoluteChromatic Mod 12) + 12) Mod 12)

End Function


' ==============================================================================
' COLOUR HELPERS
' ==============================================================================

' StringColor
' -----------
' Returns a distinct AutoCAD Color Index (ACI) for each cello string.
' The same colors are applied to string lines and, when USE_STRING_COLORS
' is True, to note circles.
'
'   stringIndex : 0 = C  1 = G  2 = D  3 = A
Private Function StringColor(stringIndex As Integer) As Long

    Select Case stringIndex
        Case 0: StringColor = acRed        ' C string – Red
        Case 1: StringColor = acGreen      ' G string – Green
        Case 2: StringColor = acCyan       ' D string – Cyan
        Case 3: StringColor = acYellow     ' A string – Yellow
        Case Else: StringColor = acWhite   ' Fallback
    End Select

End Function


' NoteColor
' ---------
' Returns a visually distinct AutoCAD Color Index (ACI) for each of the 12
' chromatic note classes. Colors are consistent across all strings and octaves,
' so the same note name always appears in the same color.
'
'   chromaticIndex : any integer; only its mod-12 value is used
Private Function NoteColor(chromaticIndex As Integer) As Long

    Select Case ((chromaticIndex Mod 12) + 12) Mod 12
        Case 0:  NoteColor = 1    ' C  – Red
        Case 1:  NoteColor = 22   ' C# – Dark orange-red
        Case 2:  NoteColor = 30   ' D  – Orange
        Case 3:  NoteColor = 50   ' D# – Yellow-green
        Case 4:  NoteColor = 2    ' E  – Yellow
        Case 5:  NoteColor = 3    ' F  – Green
        Case 6:  NoteColor = 84   ' F# – Dark green
        Case 7:  NoteColor = 4    ' G  – Cyan
        Case 8:  NoteColor = 131  ' G# – Steel blue
        Case 9:  NoteColor = 5    ' A  – Blue
        Case 10: NoteColor = 6    ' A# – Magenta
        Case 11: NoteColor = 201  ' B  – Violet
        Case Else: NoteColor = 7  '      White (fallback)
    End Select

End Function


' CircleColor
' -----------
' Dispatches to the appropriate color function based on the USE_STRING_COLORS flag.
'
'   USE_STRING_COLORS = True  → delegates to StringColor(stringIndex)
'   USE_STRING_COLORS = False → delegates to NoteColor(absoluteChromatic)
'
'   stringIndex      : 0–3 (string identity)
'   absoluteChromatic: open-string root + semitone offset (used for note-class color)
Private Function CircleColor(stringIndex As Integer, absoluteChromatic As Integer) As Long

    If USE_STRING_COLORS Then
        CircleColor = StringColor(stringIndex)
    Else
        CircleColor = NoteColor(absoluteChromatic)
    End If

End Function


' ==============================================================================
' LEGEND
' ==============================================================================

' NoteNamePT
' ----------
' Returns the note name in Brazilian Portuguese solfège notation.
' Accidentals are spelled with a sharp sign (#).
'
'   chromaticIndex : any integer; only its mod-12 value is used
Private Function NoteNamePT(chromaticIndex As Integer) As String
    Dim names(11) As String
    names(0)  = "D" & Chr(243)              ' Dó
    names(1)  = "D" & Chr(243) & "#"        ' Dó#
    names(2)  = "R" & Chr(233)              ' Ré
    names(3)  = "R" & Chr(233) & "#"        ' Ré#
    names(4)  = "Mi"                        ' Mi
    names(5)  = "F" & Chr(225)              ' Fá
    names(6)  = "F" & Chr(225) & "#"        ' Fá#
    names(7)  = "Sol"                       ' Sol
    names(8)  = "Sol#"                      ' Sol#
    names(9)  = "L" & Chr(225)              ' Lá
    names(10) = "L" & Chr(225) & "#"        ' Lá#
    names(11) = "Si"                        ' Si
    NoteNamePT = names(((chromaticIndex Mod 12) + 12) Mod 12)
End Function


' DrawColorLegend
' ---------------
' Draws a colour/note legend to the right of the fingerboard. The origin of
' the legend is LEGEND_OFFSET_X mm past the widest (bridge) edge of the diagram.
' Each row contains:
'   • A filled circle (same size and ACI colour as the fingerboard note circles)
'   • The note name in Brazilian Portuguese (LEGEND_TEXT_STYLE, LEGEND_TEXT_HEIGHT, black)
'
' Only notes that appear on the fingerboard are listed: sharps are included
' only when DRAW_SHARPS is True.
Private Sub DrawColorLegend(doc As AcadDocument)

    ' X centre of the legend circles: rightmost bridge-span edge + offset
    Dim legendX As Double
    legendX = ORIGIN_X + (NUT_STRING_SPAN + BRIDGE_STRING_SPAN) / 2# + LEGEND_OFFSET_X

    Dim row     As Integer
    Dim c       As Integer
    row = 0

    For c = 0 To 11
        If Not (IsSharpNote(c) And Not DRAW_SHARPS) Then

            Dim cy As Double
            cy = ORIGIN_Y - CDbl(row) * LEGEND_ROW_SPACING

            Dim col As Long
            col = NoteColor(c)

            ' Filled circle – identical appearance to fingerboard note circles
            Call DrawNoteCircle(doc, legendX, cy, col, LAYER_LEGEND)

            ' Note label to the right of the circle
            Dim insertPt(2) As Double
            insertPt(0) = legendX + NOTE_RADIUS + 2#
            insertPt(1) = cy - LEGEND_TEXT_HEIGHT / 2#
            insertPt(2) = 0#

            Dim txtObj As AcadText
            Set txtObj = doc.ModelSpace.AddText(NoteNamePT(c), insertPt, LEGEND_TEXT_HEIGHT)
            txtObj.Layer     = LAYER_LEGEND
            txtObj.Color     = acWhite          ' ACI 7: renders black on white paper
            txtObj.StyleName = LEGEND_TEXT_STYLE

            row = row + 1
        End If
    Next c

End Sub


' ==============================================================================
' UTILITY
' ==============================================================================

' EnsureLayer
' -----------
' Creates a layer with the given name and default color if it does not already
' exist. Leaves any existing layer with that name unchanged.
'
'   doc     : active AcadDocument object
'   lyrName : name of the layer to create / verify
'   col     : default AutoCAD Color Index applied only when the layer is created
Private Sub EnsureLayer(doc As AcadDocument, lyrName As String, col As Long)

    Dim lyr As AcadLayer
    On Error Resume Next
    Set lyr = doc.Layers.Item(lyrName)
    If Err.Number <> 0 Then
        Err.Clear
        Set lyr = doc.Layers.Add(lyrName)
        lyr.Color = col
    End If
    On Error GoTo 0

End Sub


' EnsureTextStyle
' ---------------
' Creates a text style with the given name and TrueType font file if it does
' not already exist. Leaves any existing style with that name unchanged.
'
'   doc       : active AcadDocument object
'   styleName : name of the text style to create / verify
'   fontFile  : TrueType font filename (e.g. "arial.ttf")
Private Sub EnsureTextStyle(doc As AcadDocument, styleName As String, fontFile As String)

    Dim sty As AcadTextStyle
    On Error Resume Next
    Set sty = doc.TextStyles.Item(styleName)
    If Err.Number <> 0 Then
        Err.Clear
        Set sty = doc.TextStyles.Add(styleName)
        sty.fontFile = fontFile
    End If
    On Error GoTo 0

End Sub
