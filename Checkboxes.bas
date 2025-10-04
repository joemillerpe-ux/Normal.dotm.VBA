Attribute VB_Name = "Checkboxes"
'Note: VBA uses UTF-16 hex codes with ChrW(). Lookup code and replace 0x with &H

Sub ToggleCheckBox()
    ' Start Undo block
    Application.UndoRecord.StartCustomRecord "Toggle Checkbox"

    Dim rng As Range
    Dim found As Boolean
    Dim nextChar As String
    Dim insertionStart As Long
    Set rng = Selection.Range

    ' Unicode characters for checkboxes
    Dim emptyBox As String, checkedBox As String
    emptyBox = ChrW(&H2B1C) ' ? (Empty Checkbox)
    checkedBox = ChrW(&H2705) ' ? (Checked Checkbox)

    ' Expand selection to check surrounding characters
    found = False

    ' Check if the cursor is at or next to a checkbox
    If rng.Start > 1 Then
        rng.MoveStart wdCharacter, -1
        If rng.Text = emptyBox Or rng.Text = checkedBox Then
            found = True
        Else
            rng.MoveStart wdCharacter, 1
        End If
    End If

    If rng.End < ActiveDocument.Content.End Then
        rng.MoveEnd wdCharacter, 1
        If rng.Text = emptyBox Or rng.Text = checkedBox Then
            found = True
        Else
            rng.MoveEnd wdCharacter, -1
        End If
    End If

    ' If a checkbox was found, toggle it
    If found Then
        If rng.Text = emptyBox Then
            rng.Text = checkedBox
        ElseIf rng.Text = checkedBox Then
            rng.Text = emptyBox
        End If

        ' Force font to Calibri, inherit size
        With rng.Font
            .name = "Calibri"
            .Color = wdColorAutomatic
            .Bold = False
            .Italic = False
            .Underline = wdUnderlineNone
            ' .Size is intentionally not set to inherit surrounding size
        End With
    Else
        ' Insert new checkbox
        insertionStart = rng.End
        rng.InsertAfter emptyBox
        rng.Collapse wdCollapseEnd

        If rng.End < ActiveDocument.Content.End Then
            rng.MoveEnd wdCharacter, 1
            nextChar = rng.Text
            rng.MoveEnd wdCharacter, -1
        Else
            nextChar = ""
        End If

        If nextChar <> " " Then
            rng.InsertAfter " "
            rng.Collapse wdCollapseEnd
        End If

        ' Apply formatting to inserted range (box + optional space)
        Dim fmtRng As Range
        Set fmtRng = ActiveDocument.Range(insertionStart, rng.End)
        With fmtRng.Font
            .name = "Calibri"
            .Color = wdColorAutomatic
            .Bold = False
            .Italic = False
            .Underline = wdUnderlineNone
            ' .Size not set — it inherits from surrounding text
        End With
    End If

    ' End Undo block
    Application.UndoRecord.EndCustomRecord
End Sub
Sub UncheckCheckBoxesInSelection()
    ' Start Undo block
    Application.UndoRecord.StartCustomRecord "Uncheck Selected Checkboxes"

    Dim rng As Range
    Dim checkedBox As String, emptyBox As String

    checkedBox = ChrW(&H2705) ' ? Checked Checkbox
    emptyBox = ChrW(&H2B1C)   ' ? Empty Checkbox

    Set rng = Selection.Range

    With rng.Find
        .ClearFormatting
        .Text = checkedBox
        .Replacement.ClearFormatting
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
    End With

    Do While rng.Find.Execute
        rng.Text = emptyBox
        rng.Collapse wdCollapseEnd
    Loop

    ' End Undo block
    Application.UndoRecord.EndCustomRecord
End Sub


Sub UncheckAllCheckboxes()
    ' Start Undo block
    Application.UndoRecord.StartCustomRecord "Uncheck All Checkboxes"

    Dim Doc As Document
    Dim rng As Range
    Dim checkedBox As String, emptyBox As String

    Set Doc = ActiveDocument
    checkedBox = ChrW(&H2705) ' ? Checked Checkbox
    emptyBox = ChrW(&H2B1C)   ' ? Empty Checkbox

    Set rng = Doc.Content
    With rng.Find
        .ClearFormatting
        .Text = checkedBox
        .Replacement.ClearFormatting
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
    End With

    Do While rng.Find.Execute
        rng.Text = emptyBox
        rng.Collapse wdCollapseEnd
    Loop

    ' End Undo block
    Application.UndoRecord.EndCustomRecord
End Sub




Sub RemoveCheckboxes()
    ' Start Undo block
    Application.UndoRecord.StartCustomRecord "Remove Checkboxes"
    
    Dim rng As Range
    Set rng = ActiveDocument.Content

    ' Remove Empty Checkbox (? - Unicode U+2B1C)
    With rng.Find
        .Text = ChrW(&H2B1C)
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .Execute Replace:=wdReplaceAll
    End With

    ' Remove Checked Checkbox (? - Unicode U+2705)
    With rng.Find
        .Text = ChrW(&H2705)
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .Execute Replace:=wdReplaceAll
    End With

    ' End Undo block
    Application.UndoRecord.EndCustomRecord
End Sub

Sub ToggleHeadingCollapseState()
    Dim Para As Paragraph
    Dim I As Integer
    Dim StyleName As String
    Dim collapseNow As Boolean
    Dim firstHeadingLevel As Integer
    Dim styleCollapseDefault As Boolean

    ' Start Undo record
    Application.UndoRecord.StartCustomRecord "Toggle Heading Collapse State"

    ' Find the first heading to determine current CollapseByDefault setting
    For Each Para In ActiveDocument.Paragraphs
        If Para.OutlineLevel >= wdOutlineLevel1 And Para.OutlineLevel <= wdOutlineLevel9 Then
            firstHeadingLevel = Para.OutlineLevel
            StyleName = "Heading " & firstHeadingLevel
            styleCollapseDefault = ActiveDocument.Styles(StyleName).ParagraphFormat.CollapsedByDefault
            collapseNow = Not styleCollapseDefault
            Exit For
        End If
    Next Para

    If StyleName = "" Then
        MsgBox "No standard headings (Heading 1–9) found in the document.", vbExclamation
        GoTo ExitPoint
    End If

    ' Set all paragraph CollapsedState
    For Each Para In ActiveDocument.Paragraphs
        If Para.OutlineLevel >= wdOutlineLevel1 And Para.OutlineLevel <= wdOutlineLevel9 Then
            Para.CollapsedState = collapseNow
        End If
    Next Para

    ' Set all heading styles' CollapsedByDefault
    For I = 1 To 9
        StyleName = "Heading " & I
        With ActiveDocument.Styles(StyleName).ParagraphFormat
            .CollapsedByDefault = collapseNow
        End With
    Next I

ExitPoint:
    Application.UndoRecord.EndCustomRecord
End Sub

