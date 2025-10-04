Attribute VB_Name = "MakeTextCell"
Sub ConvertSelectedTextToPlainTextControl()
'
' ConvertSelectedTextToPlainTextControl Subroutine
' Manual version: Convert selected text to Plain Text Content Control
' This is the manual version for individual text selection
'
    Dim cc As ContentControl
    Dim selectedRange As Range
    Dim Doc As Document
    Dim originalText As String
    
    Set Doc = ActiveDocument
    
    ' Check if text is selected
    If Selection.Type = wdSelectionIP Then
        Exit Sub
    End If
    
    ' Start undo record for complete operation
    Application.UndoRecord.StartCustomRecord "Convert to Plain Text Control"
    
    ' Store the selected range details
    Set selectedRange = Selection.Range.Duplicate
    originalText = Trim(selectedRange.Text)
    Dim originalStart As Long, originalEnd As Long
    originalStart = selectedRange.Start
    originalEnd = selectedRange.End
    
    ' Check if selection ends at end of paragraph and add space if needed
    If selectedRange.End = selectedRange.Paragraphs(selectedRange.Paragraphs.count).Range.End - 1 Then
        ' We're at the end of a paragraph, add a space after the selection
        Selection.Collapse wdCollapseEnd
        Selection.TypeText " "
        
        ' Reselect the original text (without the added space)
        Selection.SetRange originalStart, originalEnd
        Set selectedRange = Selection.Range.Duplicate
    End If
    
    ' Convert to Normal style before processing
    Selection.Style = ActiveDocument.Styles("Normal")  ' Convert to Normal style before processing
    
    ' Create Plain Text Content Control from selection
    Set cc = Doc.ContentControls.Add(wdContentControlText, Selection.Range)
    
    ' Enhanced properties setup
    With cc
        .title = ""  ' No visible title
        .Tag = "PlainText_" & format(Now, "yyyymmdd_hhmmss") & "_" & Left(Replace(originalText, " ", ""), 10)
        .LockContentControl = False  ' Allow control deletion
        .LockContents = False        ' Allow content editing
        .Appearance = wdContentControlBoundingBox  ' Try bounding box instead of tags
        
        ' Set the plain text content
        .Range.Text = originalText
        
        ' Format the text: Courier New 9pt dark red with light gray background
        With .Range.Font
            .name = "Courier New"
            .Size = 9
            .Color = RGB(139, 0, 0)  ' Dark red color
        End With
        
        ' Remove all paragraph spacing to minimize height
        With .Range.ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacing = LinesToPoints(1)  ' Single line spacing
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
        End With
        
        ' Add light gray background
        .Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)  ' Light gray background
        
        ' Set placeholder text (only shows when empty)
        If Len(Trim(originalText)) = 0 Then
            .PlaceholderText = "Enter plain text here..."
        End If
    End With
    
    ' Position cursor at end of the new content control
    cc.Range.Collapse wdCollapseEnd
    Selection.SetRange cc.Range.End, cc.Range.End  ' Keep cursor in place
    
    ' Optional: Add custom document property to track content controls
    On Error Resume Next
    Doc.CustomDocumentProperties("ContentControlCount").Delete
    On Error GoTo 0
    Doc.CustomDocumentProperties.Add _
        name:="ContentControlCount", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeNumber, _
        Value:=Doc.ContentControls.count
    
    ' Enhanced error handling and cleanup
    On Error GoTo ErrorHandler
    
    ' Refresh document view to ensure proper display (without moving cursor)
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    
    ' End undo record
    Application.UndoRecord.EndCustomRecord
    
    Exit Sub
    
ErrorHandler:
    ' Cleanup on error
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "Error converting text to Plain Text Control: " & Err.Description, vbExclamation
End Sub



