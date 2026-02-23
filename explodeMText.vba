Option Explicit

'================================================================================
' Module : explodeMText.vba
' Purpose: Iterate all MText objects on layers "Texto" and "Dobras", strip their
'          MText formatting codes, then explode each one into a single plain
'          Text (DText) entity.
'
' Problem solved:
'   Raw MText content such as:
'     \A1;{\fCalibri|b0|i0|c0|p34;DS 65\H0.7x;\SO^;}
'   contains two formatting "runs" (normal height text + smaller-height stacked
'   text for the degree sign).  Exploding that MText directly yields TWO separate
'   Text objects: one with "DS 65" and one with the degree character "°".
'
'   This module first rewrites the TextString to plain text (one single run),
'   so that explosion always produces exactly ONE Text object per MText.
'
' Reference: mText.vba  – handle-based _SELECT + _EXPLODE approach.
'================================================================================


'--------------------------------------------------------------------------------
' ExplodeMText_TexDobras  (entry point – call this from VBARUN)
'
' Layers processed: "Texto" and "Dobras"
'--------------------------------------------------------------------------------
Public Sub ExplodeMText_TexDobras()

    Dim acadDoc    As Object   ' AcadDocument / ThisDrawing
    Dim modelSpace As Object   ' AcadModelSpace
    Dim entity     As Object   ' generic loop iterator
    Dim mtextObj   As Object   ' current MText entity
    Dim mtextList  As Collection
    Dim i          As Long
    Dim cleanText  As String
    Dim processed  As Long

    On Error GoTo ErrorHandler

    Set acadDoc    = ThisDrawing
    Set modelSpace = acadDoc.ModelSpace

    ' ── Step 1: collect every MText on the target layers ─────────────────────
    Set mtextList = New Collection

    For Each entity In modelSpace
        If entity.ObjectName = "AcDbMText" Then
            If entity.Layer = "Texto" Or entity.Layer = "Dobras" Then
                mtextList.Add entity
            End If
        End If
    Next entity

    If mtextList.Count = 0 Then
        MsgBox "Nenhum MText encontrado nas camadas 'Texto' ou 'Dobras'.", _
               vbInformation, "ExplodeMText"
        GoTo Cleanup
    End If

    ' ── Step 2: simplify TextString then explode each MText ──────────────────
    processed = 0

    For i = 1 To mtextList.Count
        Set mtextObj = mtextList(i)

        On Error Resume Next

        ' Strip all MText formatting codes.
        ' \SO^;  →  "°"  (degree-sign idiom used in these drawings).
        ' All other codes (font, height, alignment, etc.) are removed.
        cleanText = SimplifyMTextRaw(mtextObj.TextString)

        ' Overwrite the TextString with the plain, single-run result.
        ' After this call the MText has no height changes mid-string,
        ' so explosion will yield exactly one Text entity.
        mtextObj.TextString = cleanText
        mtextObj.Update

        If Err.Number <> 0 Then
            Debug.Print "Simplify failed  handle=" & mtextObj.Handle & _
                        "  err=" & Err.Description
            Err.Clear
        Else
            ' Explode using the same handle-based pattern as mText.vba
            Dim explodeCmd As String
            explodeCmd = "_SELECT (handent """ & mtextObj.Handle & """) " & vbCr & _
                         "_EXPLODE " & vbCr
            acadDoc.SendCommand explodeCmd

            If Err.Number = 0 Then
                processed = processed + 1
            Else
                Debug.Print "Explode failed  handle=" & mtextObj.Handle & _
                            "  err=" & Err.Description
                Err.Clear
            End If
        End If

        On Error GoTo ErrorHandler
    Next i

    ' Allow BricsCAD to finish processing the last commands before refreshing
    Wait 0.5
    acadDoc.Regen acAllViewports

    MsgBox processed & " de " & mtextList.Count & _
           " objeto(s) MText convertido(s) para Text com sucesso.", _
           vbInformation, "ExplodeMText"

Cleanup:
    Set acadDoc    = Nothing
    Set modelSpace = Nothing
    Set mtextList  = Nothing
    Set entity     = Nothing
    Set mtextObj   = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erro: " & Err.Description & vbCrLf & "(Código " & Err.Number & ")", _
           vbCritical, "ExplodeMText"
    Resume Cleanup

End Sub


'================================================================================
' SimplifyMTextRaw
'
' Strips BricsCAD / AutoCAD MText inline formatting codes from a raw TextString
' and returns a clean, printable string suitable for a plain Text entity.
'
' Conversion table (applied in order):
'
'   \SO^;  or  \So^;          →  "°"   degree-sign idiom (superscript O)
'   \S…;   (any remaining)    →  ""    other stacked / fraction text removed
'   \f…;   \F…;               →  ""    font change
'   \A[n];                    →  ""    alignment
'   \H[n]x; or \H[n];         →  ""    character height multiplier
'   \W[n];                    →  ""    width factor
'   \Q[n];                    →  ""    oblique angle
'   \T[n];                    →  ""    tracking
'   \C[n];                    →  ""    colour index
'   \p…;                      →  ""    paragraph properties
'   \L \l \O \o \K \k          →  ""    under/over-line and strikethrough toggles
'   \P                         →  " "   paragraph break → space
'   %%d  %%D                   →  "°"  alternate degree control code
'   {  }                       →  ""    group delimiters
'================================================================================
Private Function SimplifyMTextRaw(ByVal rawStr As String) As String

    Dim s  As String
    Dim re As Object

    s = rawStr
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True    ' treat upper/lower the same for all letter-based codes

    ' ── Degree-sign idiom: \SO^; or \So^;  →  "°" ────────────────────────────
    ' This MUST be done before the general \S handler and before the \O toggle
    ' handler, otherwise the "O" inside \SO^; would be consumed by those steps.
    '
    ' Pattern breakdown (VBScript regex):
    '   \\S  – literal backslash + S   (matches the \S in the raw string)
    '   O    – letter O (case-insensitive: also matches o)
    '   \^   – literal caret character  (the stacking separator)
    '   [^;]* – anything up to the closing semicolon
    '   ;    – closing semicolon
    re.Pattern = "\\SO\^[^;]*;"
    s = re.Replace(s, Chr(176))         ' Chr(176) = U+00B0 DEGREE SIGN  "°"

    ' ── Remaining stacked text: \S…;  →  remove ──────────────────────────────
    ' Any \S not already replaced above (fractions, other stacks) is dropped.
    re.Pattern = "\\S[^;]+;"
    s = re.Replace(s, "")

    ' ── Font code: \fCalibri|b0|i0|c0|p34;  →  remove ───────────────────────
    re.Pattern = "\\[fF][^;]*;"
    s = re.Replace(s, "")

    ' ── Alignment: \A1;  →  remove ───────────────────────────────────────────
    re.Pattern = "\\A\d+;"
    s = re.Replace(s, "")

    ' ── Character height: \H0.7x;  or  \H2.5;  →  remove ────────────────────
    re.Pattern = "\\H[\d.]+[xX]?;"
    s = re.Replace(s, "")

    ' ── Width factor: \W1.2;  →  remove ─────────────────────────────────────
    re.Pattern = "\\W[\d.]+;"
    s = re.Replace(s, "")

    ' ── Oblique angle: \Q15;  →  remove ─────────────────────────────────────
    re.Pattern = "\\Q[\d.]+;"
    s = re.Replace(s, "")

    ' ── Tracking: \T1.2;  →  remove ─────────────────────────────────────────
    re.Pattern = "\\T[\d.]+;"
    s = re.Replace(s, "")

    ' ── Colour index: \C2;  →  remove ────────────────────────────────────────
    re.Pattern = "\\C\d+;"
    s = re.Replace(s, "")

    ' ── Paragraph properties: \pi1,l1,ql;  →  remove ────────────────────────
    re.Pattern = "\\p[^;]*;"
    s = re.Replace(s, "")

    ' ── Inline toggles (no closing semicolon): \L \l \O \o \K \k ────────────
    ' \L/\l = underline on/off,  \O/\o = overline on/off,  \K/\k = strikethrough on/off
    ' NOTE: \O toggle is safe here because \SO^; was already handled above.
    re.Pattern = "\\[LlOoKk]"
    s = re.Replace(s, "")

    ' ── Paragraph break: \P  →  space ────────────────────────────────────────
    re.Pattern = "\\P"
    s = re.Replace(s, " ")

    ' ── Legacy degree control code: %%d  →  "°" ─────────────────────────────
    s = Replace(s, "%%d", Chr(176))
    s = Replace(s, "%%D", Chr(176))

    ' ── Remove { } group delimiters ──────────────────────────────────────────
    s = Replace(s, "{", "")
    s = Replace(s, "}", "")

    ' ── Collapse any resulting runs of whitespace ─────────────────────────────
    re.Pattern = "\s+"
    s = re.Replace(s, " ")

    SimplifyMTextRaw = Trim(s)

End Function


'--------------------------------------------------------------------------------
' Wait – yield control for the given number of seconds
'--------------------------------------------------------------------------------
Private Sub Wait(seconds As Double)
    Dim t As Double
    t = Timer
    Do While Timer < t + seconds
        DoEvents
    Loop
End Sub


' ================================================================================
' HOW TO USE
' ================================================================================
' 1. Open BricsCAD and the target drawing.
' 2. Press ALT+F11 to open the VBA IDE (or type VBAIDE in the command line).
' 3. Go to File > Import File... and select this file (explodeMText.vba),
'    OR go to Insert > Module and paste this code into the new module.
' 4. Close the VBA IDE.
' 5. In BricsCAD, type VBARUN in the command line.
' 6. Select "ExplodeMText_TexDobras" from the list and click Run.
'
' WHAT IT DOES
' ─────────────
' For every MText on layers "Texto" and "Dobras" the macro will:
'   a) Read the raw TextString  (e.g.  \A1;{\fCalibri|b0|i0|c0|p34;DS 65\H0.7x;\SO^;})
'   b) Strip all formatting codes and convert the degree-sign idiom to "°"
'      → result:  DS 65°
'   c) Write the plain text back to the MText object (now a single formatting run)
'   d) Explode the MText via EXPLODE command
'      → result:  ONE regular Text entity containing "DS 65°"
' ================================================================================
