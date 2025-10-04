Attribute VB_Name = "AI_FormatAItoWord"
Option Explicit
Sub ApplyFormattingStandards()
    ' ==============================================================================
    ' USER PARAMETERS
    ' ==============================================================================
    Const CodeTableBorderWeight As Single = 0          ' Borderless table for code blocks
    Const BlockquoteIndentIncrement As Single = 0.25   ' Inches per nesting level (increase for deeper indents)
    Const ListNestingSpaces As Integer = 2             ' Number of spaces per list nesting level
    
    ' ==============================================================================
    ' MAIN PROCEDURE - INTERACTIVE STEP-BY-STEP MODE
    ' ==============================================================================
    Dim Doc As Document
    Dim Para As Paragraph
    Dim I As Long
    Dim ProcessedParas As Collection
    Dim MissingStyles As String
    Dim UserResponse As VbMsgBoxResult
    Dim DebugMode As Boolean
    Dim StartTime As Single
    Dim EndTime As Single
    Dim ElapsedSeconds As Long
    Dim MinutesStr As String
    Dim SecondsStr As String
    
    Set Doc = ActiveDocument
    Set ProcessedParas = New Collection
    
    ' Start timer
    StartTime = Timer
    
    ' Enable undo for entire operation
    Application.UndoRecord.StartCustomRecord "Apply Markdown Formatting"
    
    ' *** OPTIMIZATION #1: Disable screen updating ***
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' Ask if user wants debug mode
    UserResponse = MsgBox("Run in DEBUG MODE?" & vbCrLf & vbCrLf & _
                          "YES = Step-by-step prompts + process links and images" & vbCrLf & _
                          "NO = Automatic processing (skips links and images)" & vbCrLf & vbCrLf & _
                          "Debug mode includes Steps 6 (Links) and 7 (Images)." & vbCrLf & _
                          "Normal mode skips these steps for faster processing.", _
                          vbYesNo + vbQuestion, "Debug Mode?")
    
    DebugMode = (UserResponse = vbYes)
    
    ' ====== STEP 1: Check for required styles ======
    Application.StatusBar = "Ready: Step 1 - Check for required styles"
    DoEvents
    
    If DebugMode Then
        UserResponse = MsgBox("STEP 1: Check for required styles" & vbCrLf & vbCrLf & _
                              "This will verify that all necessary Word styles exist." & vbCrLf & vbCrLf & _
                              "Continue?", vbYesNoCancel + vbQuestion, "Step 1 of 7")
        
        If UserResponse = vbCancel Then GoTo CleanExit
        If UserResponse = vbNo Then GoTo Step2
    End If
    
    Application.StatusBar = "Processing: Step 1 - Checking styles..."
    DoEvents
    
    MissingStyles = CheckRequiredStyles(Doc)
    
    Application.StatusBar = "Completed: Step 1 - Style check complete"
    DoEvents
    
    If MissingStyles <> "" Then
        MsgBox "WARNING - Step 1: The following styles are missing:" & vbCrLf & MissingStyles, _
               vbExclamation, "Missing Styles Detected"
    End If
    
    ' Apply "No Spacing" as base style to entire document
    Application.StatusBar = "Applying base style..."
    DoEvents
    
    On Error Resume Next
    Doc.Range.Style = "No Spacing"
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Completed: Step 1 - Base style applied"
    DoEvents
    
Step2:
    ' ====== STEP 2: Process code blocks ======
    Application.StatusBar = "Ready: Step 2 - Process code blocks"
    DoEvents
    
    If DebugMode Then
        UserResponse = MsgBox("STEP 2: Process code blocks (``` fences)" & vbCrLf & vbCrLf & _
                              "This will convert fenced code blocks into borderless tables." & vbCrLf & vbCrLf & _
                              "Continue?", vbYesNoCancel + vbQuestion, "Step 2 of 7")
        
        If UserResponse = vbCancel Then GoTo CleanExit
        If UserResponse = vbNo Then GoTo Step3
    End If
    
    Application.StatusBar = "Processing: Step 2 - Converting code blocks..."
    DoEvents
    
    ProcessCodeBlocks Doc
    
    Application.StatusBar = "Completed: Step 2 - Code blocks processed"
    DoEvents
    
Step3:
    ' ====== STEP 3: Process tables (BEFORE inline formatting to protect table content) ======
    Application.StatusBar = "Ready: Step 3 - Process Markdown tables"
    DoEvents
    
    If DebugMode Then
        UserResponse = MsgBox("STEP 3: Process Markdown tables (pipe-delimited)" & vbCrLf & vbCrLf & _
                              "This will convert | table | rows | into Word tables." & vbCrLf & vbCrLf & _
                              "NOTE: Tables converted early to protect from inline formatting." & vbCrLf & vbCrLf & _
                              "Continue?", vbYesNoCancel + vbQuestion, "Step 3 of 7")
        
        If UserResponse = vbCancel Then GoTo CleanExit
        If UserResponse = vbNo Then GoTo Step4
    End If
    
    Application.StatusBar = "Processing: Step 3 - Converting Markdown tables..."
    DoEvents
    
    ProcessMarkdownTables Doc
    
    Application.StatusBar = "Completed: Step 3 - Tables processed"
    DoEvents
    
Step4:
    ' ====== STEP 4: Process block elements ======
    ' Two-pass approach: non-lists backward (avoids shifts), lists forward (builds hierarchy)
    Application.StatusBar = "Ready: Step 4 - Process block elements"
    DoEvents
    
    If DebugMode Then
        UserResponse = MsgBox("STEP 4: Process block elements" & vbCrLf & vbCrLf & _
                              "This will process:" & vbCrLf & _
                              "  • Headings (# ## ###)" & vbCrLf & _
                              "  • Horizontal rules (---)" & vbCrLf & _
                              "  • Blockquotes (>)" & vbCrLf & _
                              "  • Lists (- * 1. 1))" & vbCrLf & vbCrLf & _
                              "Continue?", vbYesNoCancel + vbQuestion, "Step 4 of 7")
        
        If UserResponse = vbCancel Then GoTo CleanExit
        If UserResponse = vbNo Then GoTo Step5
    End If
    
    Application.StatusBar = "Processing: Step 4 - Processing block elements..."
    DoEvents
    
    Dim TotalParas As Long
    TotalParas = Doc.Paragraphs.count
    
    ' PASS 1: Process non-list elements backward (avoids paragraph shifts)
    For I = Doc.Paragraphs.count To 1 Step -1
        If I <= Doc.Paragraphs.count Then
            ' Update status every 10 paragraphs
            If I Mod 10 = 0 Then
                Application.StatusBar = "Processing block elements (pass 1): paragraph " & (TotalParas - I + 1) & " of " & TotalParas
                DoEvents
            End If
            
            Set Para = Doc.Paragraphs(I)
            
            ' Skip table cells (both code blocks and markdown tables)
            If Para.Range.Tables.count > 0 Then
                GoTo NextPara4Pass1
            End If
            
            ' Process only non-list items in this pass
            If ProcessHeading(Para) Then GoTo NextPara4Pass1
            If ProcessHorizontalRule(Para) Then GoTo NextPara4Pass1
            If ProcessBlockquote(Para) Then GoTo NextPara4Pass1
            
            ' Apply No Spacing style to non-empty, non-list paragraphs
            Dim IsListItem As Boolean
            Dim TempText As String
            TempText = Trim(Para.Range.Text)
            
            ' Quick check if this looks like a list item (don't process yet)
            IsListItem = (Left(TempText, 2) = "- " Or Left(TempText, 2) = "* " Or IsNumberedListItem(TempText))
            
            If Para.Range.Text <> vbCr And Not IsListItem Then
                Para.Style = "No Spacing"
            End If
            
NextPara4Pass1:
        End If
    Next I
    
    Application.StatusBar = "Processing lists (pass 2)..."
    DoEvents
    
    ' PASS 2: Process lists FORWARD (builds proper Level 1 ? Level 2 hierarchy)
    For I = 1 To Doc.Paragraphs.count
        If I <= Doc.Paragraphs.count Then
            ' Update status every 10 paragraphs
            If I Mod 10 = 0 Then
                Application.StatusBar = "Processing lists (pass 2): paragraph " & I & " of " & Doc.Paragraphs.count
                DoEvents
            End If
            
            Set Para = Doc.Paragraphs(I)
            
            ' Skip table cells
            If Para.Range.Tables.count > 0 Then
                GoTo NextPara4Pass2
            End If
            
            ' Process lists with index for smart continuation
            ProcessList Para, I
            
NextPara4Pass2:
        End If
    Next I
    
    Application.StatusBar = "Completed: Step 4 - Block elements processed"
    DoEvents
    
Step5:
    ' ====== STEP 5: Process inline formatting ======
    Application.StatusBar = "Ready: Step 5 - Process inline formatting"
    DoEvents
    
    If DebugMode Then
        UserResponse = MsgBox("STEP 5: Process inline formatting" & vbCrLf & vbCrLf & _
                              "This will process:" & vbCrLf & _
                              "  • Bold (**text**)" & vbCrLf & _
                              "  • Italic (*text*)" & vbCrLf & _
                              "  • Strikethrough (~~text~~)" & vbCrLf & _
                              "  • Inline code (`code`)" & vbCrLf & vbCrLf & _
                              "Note: Will process table cells too!" & vbCrLf & vbCrLf & _
                              "Continue?", vbYesNoCancel + vbQuestion, "Step 5 of 7")
        
        If UserResponse = vbCancel Then GoTo CleanExit
        If UserResponse = vbNo Then GoTo Step6
    End If
    
    Application.StatusBar = "Processing: Step 5 - Applying inline formatting..."
    DoEvents
    
    ' Process with error handling for each paragraph
    On Error GoTo Step5Error
    Dim TotalParas5 As Long
    TotalParas5 = Doc.Paragraphs.count
    
    For I = 1 To Doc.Paragraphs.count
        ' Validate paragraph exists
        If I > Doc.Paragraphs.count Then Exit For
        
        ' Update status every 10 paragraphs
        If I Mod 10 = 0 Then
            Application.StatusBar = "Processing inline formatting: paragraph " & I & " of " & TotalParas5
            DoEvents
        End If
        
        Set Para = Doc.Paragraphs(I)
        
        ' Skip code block table cells (they have "Code" style applied)
        If Para.Range.Tables.count > 0 Then
            On Error Resume Next
            If Para.Style = "Code" Or Para.Range.Style = "Code" Then
                On Error GoTo Step5Error
                GoTo NextPara5
            End If
            On Error GoTo Step5Error
        End If
        
        ' Skip empty paragraphs
        If Len(Trim(Para.Range.Text)) > 1 Then
            ProcessInlineFormatting Para
        End If
        
NextPara5:
    Next I
    
    On Error GoTo 0
    
    Application.StatusBar = "Completed: Step 5 - Inline formatting processed"
    DoEvents
    
    GoTo Step6
    
Step5Error:
    MsgBox "ERROR in STEP 5 (Inline Formatting) at paragraph " & I & " of " & Doc.Paragraphs.count & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Paragraph text preview: " & Left(Para.Range.Text, 100), _
           vbCritical, "Step 5 Error"
    
    UserResponse = MsgBox("Continue processing remaining paragraphs?", vbYesNo + vbQuestion, "Continue?")
    If UserResponse = vbYes Then
        Resume Next
    Else
        GoTo CleanExit
    End If
    
Step6:
    ' ====== STEP 6: Process links (DEBUG MODE ONLY) ======
    If Not DebugMode Then GoTo Finished
    
    Application.StatusBar = "Ready: Step 6 - Process links"
    DoEvents
    
    UserResponse = MsgBox("STEP 6: Process links" & vbCrLf & vbCrLf & _
                          "This will process:" & vbCrLf & _
                          "  • Markdown links ([text](url))" & vbCrLf & _
                          "  • Bare URLs (http://...)" & vbCrLf & vbCrLf & _
                          "Continue?", vbYesNoCancel + vbQuestion, "Step 6 of 7")
    
    If UserResponse = vbCancel Then GoTo CleanExit
    If UserResponse = vbNo Then GoTo Step7
    
    Application.StatusBar = "Processing: Step 6 - Creating hyperlinks..."
    DoEvents
    
    On Error GoTo Step6Error
    Dim TotalParas6 As Long
    TotalParas6 = Doc.Paragraphs.count
    
    For I = 1 To Doc.Paragraphs.count
        If I > Doc.Paragraphs.count Then Exit For
        
        ' Update status every 10 paragraphs
        If I Mod 10 = 0 Then
            Application.StatusBar = "Processing links: paragraph " & I & " of " & TotalParas6
            DoEvents
        End If
        
        Set Para = Doc.Paragraphs(I)
        
        ' Skip code block table cells (they have "Code" style applied)
        If Para.Range.Tables.count > 0 Then
            On Error Resume Next
            If Para.Style = "Code" Or Para.Range.Style = "Code" Then
                On Error GoTo Step6Error
                GoTo NextPara6
            End If
            On Error GoTo Step6Error
        End If
        
        ProcessLinks Para
        ProcessBareUrls Para
        
NextPara6:
    Next I
    
    On Error GoTo 0
    
    Application.StatusBar = "Completed: Step 6 - Links processed"
    DoEvents
    
    GoTo Step7
    
Step6Error:
    MsgBox "ERROR in STEP 6 (Links) at paragraph " & I & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Step 6 Error"
    
    UserResponse = MsgBox("Continue processing remaining paragraphs?", vbYesNo + vbQuestion, "Continue?")
    If UserResponse = vbYes Then
        Resume Next
    Else
        GoTo CleanExit
    End If
    
Step7:
    ' ====== STEP 7: Process images (DEBUG MODE ONLY) ======
    If Not DebugMode Then GoTo Finished
    
    Application.StatusBar = "Ready: Step 7 - Process images"
    DoEvents
    
    UserResponse = MsgBox("STEP 7: Process images" & vbCrLf & vbCrLf & _
                          "This will embed images from ![alt](url) syntax." & vbCrLf & vbCrLf & _
                          "Note: This may take time if downloading from URLs." & vbCrLf & vbCrLf & _
                          "Continue?", vbYesNoCancel + vbQuestion, "Step 7 of 7")
    
    If UserResponse = vbCancel Then GoTo CleanExit
    If UserResponse = vbNo Then GoTo Finished
    
    Application.StatusBar = "Processing: Step 7 - Embedding images..."
    DoEvents
    
    ProcessImages Doc
    
    Application.StatusBar = "Completed: Step 7 - Images processed"
    DoEvents
    
Finished:
    ' *** OPTIMIZATION #1: Re-enable screen updating ***
    Application.ScreenUpdating = True
    
    Application.StatusBar = False  ' Reset status bar
    Application.UndoRecord.EndCustomRecord
    
    ' Calculate elapsed time
    EndTime = Timer
    ElapsedSeconds = CLng(EndTime - StartTime)
    
    ' Handle midnight rollover
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400
    
    ' Format time as Xm Ys
    If ElapsedSeconds >= 60 Then
        MinutesStr = CStr(ElapsedSeconds \ 60) & "m "
        SecondsStr = CStr(ElapsedSeconds Mod 60) & "s"
    Else
        MinutesStr = ""
        SecondsStr = CStr(ElapsedSeconds) & "s"
    End If
    
    ' Show completion message with elapsed time
    MsgBox "Markdown formatting complete!" & vbCrLf & vbCrLf & _
           "All processing has finished successfully." & vbCrLf & vbCrLf & _
           "Time elapsed: " & MinutesStr & SecondsStr, _
           vbInformation, "Process Complete"
    
    Exit Sub
    
CleanExit:
    ' *** OPTIMIZATION #1: Re-enable screen updating ***
    Application.ScreenUpdating = True
    
    Application.StatusBar = False  ' Reset status bar
    Application.UndoRecord.EndCustomRecord
    MsgBox "Process cancelled by user.", vbInformation, "Cancelled"
    Exit Sub
    
ErrorHandler:
    ' *** OPTIMIZATION #1: Re-enable screen updating ***
    Application.ScreenUpdating = True
    
    Application.StatusBar = False  ' Reset status bar
    Application.UndoRecord.EndCustomRecord
    MsgBox "Error in Step: " & Err.Description & vbCrLf & vbCrLf & _
           "Use Ctrl+Z to undo partial changes.", vbCritical, "Error"
End Sub
Function CheckRequiredStyles(Doc As Document) As String
    ' Check if required styles exist in document
    Dim StyleNames As Variant
    Dim StyleName As Variant
    Dim MissingList As String
    Dim StyleExists As Boolean
    
    StyleNames = Array("Title", "Heading 1", "Heading 2", "Heading 3", _
                       "Normal", "Code", "Quote", "DW Array")
    
    For Each StyleName In StyleNames
        StyleExists = False
        On Error Resume Next
        StyleExists = (Doc.Styles(CStr(StyleName)).NameLocal <> "")
        On Error GoTo 0
        
        If Not StyleExists And StyleName <> "DW Array" Then
            MissingList = MissingList & "  • " & StyleName & vbCrLf
        End If
    Next StyleName
    
    CheckRequiredStyles = MissingList
End Function

Sub ProcessCodeBlocks(Doc As Document)
    ' Process code blocks delimited by ``` fences
    Dim Para As Paragraph
    Dim I As Long
    Dim InCodeBlock As Boolean
    Dim CodeBlockStart As Long
    Dim CodeBlockEnd As Long
    Dim CodeText As String
    Dim Tbl As Table
    Dim Cell As Cell
    Dim BlockCount As Integer
    
    InCodeBlock = False
    BlockCount = 0
    I = 1
    
    Application.StatusBar = "Scanning for code blocks..."
    
    Do While I <= Doc.Paragraphs.count
        Set Para = Doc.Paragraphs(I)
        
        ' Check for code fence (```)
        If Left(Trim(Para.Range.Text), 3) = "```" Then
            If Not InCodeBlock Then
                ' Start of code block
                InCodeBlock = True
                CodeBlockStart = I
                CodeText = ""
                BlockCount = BlockCount + 1
                Application.StatusBar = "Found code block #" & BlockCount & " at paragraph " & I
            Else
                ' End of code block
                InCodeBlock = False
                CodeBlockEnd = I
                
                Application.StatusBar = "Converting code block #" & BlockCount & "..."
                
                ' Extract code content between fences
                CodeText = ExtractCodeBlockContent(Doc, CodeBlockStart, CodeBlockEnd)
                
                ' Delete all paragraphs in code block
                If CodeBlockEnd <= Doc.Paragraphs.count Then
                    Doc.Range(Doc.Paragraphs(CodeBlockStart).Range.Start, _
                              Doc.Paragraphs(CodeBlockEnd).Range.End).Delete
                End If
                
                ' Insert table at position
                If CodeBlockStart <= Doc.Paragraphs.count Then
                    Set Tbl = Doc.Tables.Add(Doc.Paragraphs(CodeBlockStart).Range, 1, 1)
                Else
                    Doc.Content.InsertAfter vbCr
                    Set Tbl = Doc.Tables.Add(Doc.Range(Doc.Content.End - 1, Doc.Content.End), 1, 1)
                End If
                
                ' Format table
                With Tbl
                    .Borders.Enable = False
                    .AutoFitBehavior wdAutoFitContent
                    .PreferredWidthType = wdPreferredWidthPercent
                    .PreferredWidth = 100
                    
                    ' Add code text to cell
                    .Cell(1, 1).Range.Text = CodeText
                    .Cell(1, 1).Range.Style = "Code"
                End With
                
                ' Reset counter to continue after table
                I = CodeBlockStart
            End If
        End If
        
        I = I + 1
    Loop
    
    Application.StatusBar = "Completed: " & BlockCount & " code block(s) processed"
End Sub

Function ExtractCodeBlockContent(Doc As Document, StartPara As Long, EndPara As Long) As String
    ' Extract text between code fences, excluding the fences themselves
    Dim I As Long
    Dim Content As String
    
    For I = StartPara + 1 To EndPara - 1
        If I <= Doc.Paragraphs.count Then
            Content = Content & Doc.Paragraphs(I).Range.Text
        End If
    Next I
    
    ' Remove trailing paragraph mark
    If Right(Content, 1) = vbCr Then
        Content = Left(Content, Len(Content) - 1)
    End If
    
    ExtractCodeBlockContent = Content
End Function

Sub ProcessMarkdownTables(Doc As Document)
    ' Process Markdown tables (pipe-delimited)
    Dim Para As Paragraph
    Dim I As Long
    Dim TableStart As Long
    Dim TableEnd As Long
    Dim InTable As Boolean
    Dim TableRows As Collection
    Dim RowText As String
    Dim TableCount As Integer
    
    InTable = False
    TableCount = 0
    I = 1
    
    Application.StatusBar = "Scanning for Markdown tables..."
    
    Do While I <= Doc.Paragraphs.count
        Set Para = Doc.Paragraphs(I)
        RowText = Trim(Para.Range.Text)
        
        ' Remove paragraph mark for checking
        If Right(RowText, 1) = vbCr Then
            RowText = Left(RowText, Len(RowText) - 1)
        End If
        
        ' Check if paragraph contains pipe-delimited table row
        If InStr(RowText, "|") > 0 And Left(RowText, 1) <> ">" Then
            
            If Not InTable Then
                ' Start of table
                InTable = True
                TableStart = I
                Set TableRows = New Collection
                TableCount = TableCount + 1
                Application.StatusBar = "Found table #" & TableCount & " at paragraph " & I
            End If
            
            ' Skip alignment rows (contains only |, -, :, spaces)
            If Not IsAlignmentRow(RowText) Then
                TableRows.Add RowText
            End If
            
        ElseIf InTable Then
            ' End of table
            InTable = False
            TableEnd = I - 1
            
            Application.StatusBar = "Converting table #" & TableCount & " (" & TableRows.count & " rows)..."
            
            ' Convert to Word table
            On Error Resume Next
            ConvertToWordTable Doc, TableStart, TableEnd, TableRows
            On Error GoTo 0
            
            ' Reset counter
            I = TableStart
        End If
        
        I = I + 1
    Loop
    
    Application.StatusBar = "Completed: " & TableCount & " table(s) processed"
End Sub

Function IsAlignmentRow(RowText As String) As Boolean
    ' Check if row is alignment marker (e.g., |---|---|)
    Dim CleanText As String
    Dim Ch As String
    Dim j As Long
    
    CleanText = Replace(RowText, " ", "")
    CleanText = Replace(CleanText, vbCr, "")
    
    For j = 1 To Len(CleanText)
        Ch = Mid(CleanText, j, 1)
        If Ch <> "|" And Ch <> "-" And Ch <> ":" Then
            IsAlignmentRow = False
            Exit Function
        End If
    Next j
    
    IsAlignmentRow = True
End Function

Sub ConvertToWordTable(Doc As Document, StartPara As Long, EndPara As Long, TableRows As Collection)
    ' Convert Markdown table to Word table
    Dim Tbl As Table
    Dim RowData As Variant
    Dim Cells As Variant
    Dim I As Long, j As Long
    Dim NumRows As Long
    Dim NumCols As Long
    Dim CellText As Variant
    
    NumRows = TableRows.count
    If NumRows = 0 Then Exit Sub
    
    ' Determine number of columns from first row
    RowData = Split(TableRows(1), "|")
    NumCols = 0
    For I = LBound(RowData) To UBound(RowData)
        If Trim(RowData(I)) <> "" Then NumCols = NumCols + 1
    Next I
    
    If NumCols = 0 Then Exit Sub
    
    ' Delete original paragraphs
    On Error Resume Next
    Doc.Range(Doc.Paragraphs(StartPara).Range.Start, _
              Doc.Paragraphs(EndPara).Range.End).Delete
    On Error GoTo 0
    
    ' Insert table
    On Error Resume Next
    If StartPara > Doc.Paragraphs.count Then
        StartPara = Doc.Paragraphs.count
    End If
    
    Set Tbl = Doc.Tables.Add(Doc.Paragraphs(StartPara).Range, NumRows, NumCols)
    On Error GoTo 0
    
    If Tbl Is Nothing Then Exit Sub
    
    ' Populate table
    On Error Resume Next
    For I = 1 To NumRows
        Cells = Split(TableRows(I), "|")
        j = 1
        
        For Each CellText In Cells
            CellText = Trim(CellText)
            If CellText <> "" And j <= NumCols Then
                Tbl.Cell(I, j).Range.Text = CellText
                j = j + 1
            End If
        Next CellText
    Next I
    On Error GoTo 0
    
    ' Apply table style if exists
    On Error Resume Next
    Tbl.Style = "DW Array"
    On Error GoTo 0
End Sub
Function ProcessHeading(Para As Paragraph) As Boolean
    ' Process Markdown headings (#, ##, ###, etc.)
    Dim Text As String
    Dim Level As Integer
    Dim I As Integer
    
    Text = Para.Range.Text
    
    ' Remove paragraph mark from text for processing
    If Right(Text, 1) = vbCr Then
        Text = Left(Text, Len(Text) - 1)
    End If
    
    Level = 0
    
    ' Count leading # characters
    For I = 1 To Len(Text)
        If Mid(Text, I, 1) = "#" Then
            Level = Level + 1
        Else
            Exit For
        End If
    Next I
    
    ' Must have space after # and at least one # to be a heading
    If Level > 0 And Level <= 8 And Mid(Text, Level + 1, 1) = " " Then
        ' Extract heading text after markers
        Dim HeadingText As String
        HeadingText = Trim(Mid(Text, Level + 2))
        
        ' Replace text but preserve paragraph mark
        Dim rng As Range
        Set rng = Para.Range
        rng.End = rng.End - 1  ' Exclude paragraph mark
        rng.Text = HeadingText
        
        ' Apply style
        If Level = 1 Then
            Para.Style = "Title"
        Else
            Para.Style = "Heading " & Level - 1
        End If
        
        ProcessHeading = True
    Else
        ProcessHeading = False
    End If
End Function
Function ProcessHorizontalRule(Para As Paragraph) As Boolean
    ' Process horizontal rules (---, ***, ___) - replace with Word's standard horizontal line
    Dim Text As String
    
    Text = Para.Range.Text
    
    ' Remove paragraph mark from text for processing
    If Right(Text, 1) = vbCr Then
        Text = Left(Text, Len(Text) - 1)
    End If
    
    Text = Trim(Text)
    
    ' Check if this is a horizontal rule
    If Text = "---" Or Text = "***" Or Text = "___" Then
        ' Clear the text but preserve paragraph
        Dim rng As Range
        Set rng = Para.Range
        rng.End = rng.End - 1  ' Exclude paragraph mark
        rng.Text = ""          ' Empty the content
        
        ' Insert Word's standard horizontal line
        rng.InlineShapes.AddHorizontalLineStandard
        
        ProcessHorizontalRule = True
    Else
        ProcessHorizontalRule = False
    End If
End Function
Function ProcessBlockquote(Para As Paragraph) As Boolean
    ' Process blockquotes (>, >>, etc.)
    Dim Text As String
    Dim Level As Integer
    Dim I As Integer
    
    Text = Para.Range.Text
    
    ' Remove paragraph mark from text for processing
    If Right(Text, 1) = vbCr Then
        Text = Left(Text, Len(Text) - 1)
    End If
    
    Level = 0
    
    ' Count leading > characters
    For I = 1 To Len(Text)
        If Mid(Text, I, 1) = ">" Then
            Level = Level + 1
        ElseIf Mid(Text, I, 1) = " " Then
            ' Skip spaces between >
        Else
            Exit For
        End If
    Next I
    
    If Level > 0 Then
        ' Extract blockquote text after markers
        Dim QuoteText As String
        QuoteText = Trim(Mid(Text, I))
        
        ' Replace text but preserve paragraph mark
        Dim rng As Range
        Set rng = Para.Range
        rng.End = rng.End - 1  ' Exclude paragraph mark
        rng.Text = QuoteText
        
        ' Apply Quote style with indent
        Para.Style = "Quote"
        Para.LeftIndent = Level * InchesToPoints(0.25)
        
        ProcessBlockquote = True
    Else
        ProcessBlockquote = False
    End If
End Function
Function ProcessList(Para As Paragraph, ParaIndex As Long) As Boolean
    ' Process bulleted and numbered lists
    ' Detects list type, calculates nesting level, and applies appropriate formatting
    Dim Text As String
    Dim LeadingSpaces As Integer
    Dim Level As Integer
    
    Text = Para.Range.Text
    
    ' Remove paragraph mark from text for processing
    If Right(Text, 1) = vbCr Then
        Text = Left(Text, Len(Text) - 1)
    End If
    
    LeadingSpaces = 0
    
    ' Count leading spaces for nesting level
    For LeadingSpaces = 1 To Len(Text)
        If Mid(Text, LeadingSpaces, 1) <> " " Then Exit For
    Next LeadingSpaces
    LeadingSpaces = LeadingSpaces - 1
    
    ' Calculate nesting level (2 spaces per level)
    Level = Int(LeadingSpaces / 2) + 1
    
    Text = Trim(Text)
    
    ' Check for bullet list (-, *)
    If (Left(Text, 2) = "- " Or Left(Text, 2) = "* ") Then
        ' Extract text after marker
        Dim CleanText As String
        CleanText = Trim(Mid(Text, 3))
        
        ' Replace text but preserve paragraph mark
        Dim rng As Range
        Set rng = Para.Range
        rng.End = rng.End - 1  ' Exclude paragraph mark
        rng.Text = CleanText
        
        ApplyBulletListStyle Para, Level, ParaIndex
        ProcessList = True
        
    ' Check for numbered list (1., 1), etc.)
    ElseIf IsNumberedListItem(Text) Then
        ' Extract text after marker
        Dim CleanText2 As String
        CleanText2 = Trim(Mid(Text, InStr(Text, " ")))
        
        ' Replace text but preserve paragraph mark
        Dim rng2 As Range
        Set rng2 = Para.Range
        rng2.End = rng2.End - 1  ' Exclude paragraph mark
        rng2.Text = CleanText2
        
        ApplyNumberedListStyle Para, Level, ParaIndex
        ProcessList = True
        
    Else
        ProcessList = False
    End If
End Function
Function IsNumberedListItem(Text As String) As Boolean
    ' Check if text starts with number followed by . or )
    Dim I As Integer
    Dim FoundDigit As Boolean
    
    FoundDigit = False
    
    For I = 1 To Len(Text)
        If IsNumeric(Mid(Text, I, 1)) Then
            FoundDigit = True
        ElseIf FoundDigit And (Mid(Text, I, 1) = "." Or Mid(Text, I, 1) = ")") Then
            If Mid(Text, I + 1, 1) = " " Then
                IsNumberedListItem = True
                Exit Function
            End If
        Else
            Exit For
        End If
    Next I
    
    IsNumberedListItem = False
End Function
Sub ApplyBulletListStyle(Para As Paragraph, Level As Integer, ParaIndex As Long)
    ' Apply JDM Bullets list template with smart list continuation
    ' Creates multi-level hierarchy by continuing existing lists
    Dim ListTemplate As ListTemplate
    Dim TemplateExists As Boolean
    Dim ContinueList As Boolean
    Dim PrevPara As Paragraph
    
    On Error GoTo BulletStyleError
    
    ' Check if template exists at index 1
    TemplateExists = False
    On Error Resume Next
    If ListGalleries(wdOutlineNumberGallery).ListTemplates.count >= 1 Then
        Set ListTemplate = ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
        If Not ListTemplate Is Nothing Then
            If ListTemplate.name = "JDM Bullets" Or _
               ListTemplate.ListLevels(1).NumberStyle = wdListNumberStyleBullet Then
                TemplateExists = True
            End If
        End If
    End If
    On Error GoTo BulletStyleError
    
    ' Create template if missing
    If Not TemplateExists Then
        Dim OriginalRange As Range
        Set OriginalRange = Selection.Range.Duplicate
        
        Para.Range.Select
        InsertMyBulletList
        OriginalRange.Select
        
        Set ListTemplate = ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
    End If
    
    If ListTemplate Is Nothing Then
        Err.Raise vbObjectError + 1, , "Could not find or create 'JDM Bullets' template"
    End If
    
    ' Smart list continuation: check if previous paragraph is also a bullet list
    ContinueList = False
    
    On Error Resume Next
    If ParaIndex > 1 Then
        Set PrevPara = Para.Range.Document.Paragraphs(ParaIndex - 1)
        
        ' Continue if previous paragraph has bullet list formatting
        ' Check for both standard bullets (type 1) and outline-numbered bullets (type 4)
        If PrevPara.Range.ListFormat.ListType = wdListBullet Or _
           PrevPara.Range.ListFormat.ListType = 4 Then
            ContinueList = True
        End If
    End If
    On Error GoTo BulletStyleError
    
    ' Apply the list template
    Para.Range.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=ListTemplate, _
        ContinuePreviousList:=ContinueList, _
        ApplyTo:=wdListApplyToSelection, _
        ApplyLevel:=Level
    
    Exit Sub
    
BulletStyleError:
    MsgBox "ERROR in STEP 4 (Block Elements - Lists)" & vbCrLf & vbCrLf & _
           "Failed to apply 'JDM Bullets' list template." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure the 'InsertMyBulletList' subroutine exists.", _
           vbCritical, "List Template Error"
    Err.Raise Err.Number
End Sub
Sub ApplyNumberedListStyle(Para As Paragraph, Level As Integer, ParaIndex As Long)
    ' Apply JDM numbered list template with smart list continuation
    ' Creates multi-level hierarchy by continuing existing lists
    Dim ListTemplate As ListTemplate
    Dim TemplateExists As Boolean
    Dim ContinueList As Boolean
    Dim PrevPara As Paragraph
    
    On Error GoTo NumberStyleError
    
    ' Check if template exists at index 2
    TemplateExists = False
    On Error Resume Next
    Set ListTemplate = ListGalleries(wdOutlineNumberGallery).ListTemplates(2)
    If Not ListTemplate Is Nothing Then
        If InStr(ListTemplate.ListLevels(1).NumberFormat, "%1)") > 0 Then
            TemplateExists = True
        End If
    End If
    On Error GoTo NumberStyleError
    
    ' Create template if missing
    If Not TemplateExists Then
        Dim OriginalRange As Range
        Set OriginalRange = Selection.Range.Duplicate
        
        Para.Range.Select
        InsertJDMnumList
        OriginalRange.Select
        
        Set ListTemplate = ListGalleries(wdOutlineNumberGallery).ListTemplates(2)
    End If
    
    ' Smart list continuation: check if previous paragraph is also a numbered list
    ContinueList = False
    
    On Error Resume Next
    If ParaIndex > 1 Then
        Set PrevPara = Para.Range.Document.Paragraphs(ParaIndex - 1)
        
        ' Continue if previous paragraph has numbered list formatting
        If PrevPara.Range.ListFormat.ListType = wdListSimpleNumbering Or _
           PrevPara.Range.ListFormat.ListType = wdListOutlineNumbering Or _
           PrevPara.Range.ListFormat.ListType = wdListListNumOnly Or _
           PrevPara.Range.ListFormat.ListType = 4 Then
            ContinueList = True
        End If
    End If
    On Error GoTo NumberStyleError
    
    ' Apply the list template
    Para.Range.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=ListTemplate, _
        ContinuePreviousList:=ContinueList, _
        ApplyTo:=wdListApplyToSelection, _
        ApplyLevel:=Level
    
    Exit Sub
    
NumberStyleError:
    MsgBox "ERROR in STEP 4 (Block Elements - Lists)" & vbCrLf & vbCrLf & _
           "Failed to apply numbered list template." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure the 'InsertJDMnumList' subroutine exists.", _
           vbCritical, "List Template Error"
    Err.Raise Err.Number
End Sub
Sub ProcessInlineFormatting(Para As Paragraph)
    ' Process inline formatting: bold, italic, strikethrough, code (NOT links - handled separately)
    ' Added error handling for each step
    
    On Error Resume Next
    
    ' Process in specific order to avoid conflicts
    ProcessStrikethrough Para
    If Err.Number <> 0 Then
        Debug.Print "Error in ProcessStrikethrough: " & Err.Description
        Err.Clear
    End If
    
    ProcessBoldItalic Para
    If Err.Number <> 0 Then
        Debug.Print "Error in ProcessBoldItalic: " & Err.Description
        Err.Clear
    End If
    
    ProcessBold Para
    If Err.Number <> 0 Then
        Debug.Print "Error in ProcessBold: " & Err.Description
        Err.Clear
    End If
    
    ProcessItalic Para
    If Err.Number <> 0 Then
        Debug.Print "Error in ProcessItalic: " & Err.Description
        Err.Clear
    End If
    
    ProcessInlineCode Para
    If Err.Number <> 0 Then
        Debug.Print "Error in ProcessInlineCode: " & Err.Description
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

Sub ProcessStrikethrough(Para As Paragraph)
    ' Process ~~text~~ for strikethrough
    Dim rng As Range
    Dim StartPos As Long
    Dim EndPos As Long
    Dim TextToSearch As String
    
    Set rng = Para.Range
    
    ' Don't include the paragraph mark in our search
    TextToSearch = rng.Text
    If Right(TextToSearch, 1) = vbCr Then
        TextToSearch = Left(TextToSearch, Len(TextToSearch) - 1)
    End If
    
    Do
        StartPos = InStr(TextToSearch, "~~")
        If StartPos = 0 Then Exit Do
        
        EndPos = InStr(StartPos + 2, TextToSearch, "~~")
        If EndPos = 0 Then Exit Do
        
        ' Apply strikethrough - but only to the actual paragraph range, not including vbCr
        Set rng = Para.Range
        If Right(rng.Text, 1) = vbCr Then
            rng.End = rng.End - 1  ' Don't include paragraph mark
        End If
        
        rng.Start = rng.Start + StartPos - 1
        rng.End = rng.Start + (EndPos - StartPos) + 2
        
        ' Remove markers and format
        rng.Text = Mid(rng.Text, 3, Len(rng.Text) - 4)
        rng.Font.StrikeThrough = True
        
        ' Refresh our search text
        Set rng = Para.Range
        TextToSearch = rng.Text
        If Right(TextToSearch, 1) = vbCr Then
            TextToSearch = Left(TextToSearch, Len(TextToSearch) - 1)
        End If
    Loop
End Sub

Sub ProcessBoldItalic(Para As Paragraph)
    ' Process ***text*** or ___text___ for bold+italic
    ProcessBoldItalicMarker Para, "***"
    ProcessBoldItalicMarker Para, "___"
End Sub

Sub ProcessBoldItalicMarker(Para As Paragraph, Marker As String)
    Dim rng As Range
    Dim StartPos As Long
    Dim EndPos As Long
    Dim TextToSearch As String
    
    Set rng = Para.Range
    
    ' Don't include the paragraph mark in our search
    TextToSearch = rng.Text
    If Right(TextToSearch, 1) = vbCr Then
        TextToSearch = Left(TextToSearch, Len(TextToSearch) - 1)
    End If
    
    Do
        StartPos = InStr(TextToSearch, Marker)
        If StartPos = 0 Then Exit Do
        
        EndPos = InStr(StartPos + Len(Marker), TextToSearch, Marker)
        If EndPos = 0 Then Exit Do
        
        ' Apply bold+italic - but only to the actual paragraph range, not including vbCr
        Set rng = Para.Range
        If Right(rng.Text, 1) = vbCr Then
            rng.End = rng.End - 1  ' Don't include paragraph mark
        End If
        
        rng.Start = rng.Start + StartPos - 1
        rng.End = rng.Start + (EndPos - StartPos) + Len(Marker)
        
        ' Remove markers and format
        rng.Text = Mid(rng.Text, Len(Marker) + 1, Len(rng.Text) - Len(Marker) * 2)
        rng.Font.Bold = True
        rng.Font.Italic = True
        
        ' Refresh our search text
        Set rng = Para.Range
        TextToSearch = rng.Text
        If Right(TextToSearch, 1) = vbCr Then
            TextToSearch = Left(TextToSearch, Len(TextToSearch) - 1)
        End If
    Loop
End Sub

Sub ProcessBold(Para As Paragraph)
    ' Process **text** for bold (but not *** which is handled by ProcessBoldItalic)
    Dim rng As Range
    Dim StartPos As Long
    Dim EndPos As Long
    Dim TextToSearch As String
    
    Set rng = Para.Range
    
    ' Don't include the paragraph mark in our search
    TextToSearch = rng.Text
    If Right(TextToSearch, 1) = vbCr Then
        TextToSearch = Left(TextToSearch, Len(TextToSearch) - 1)
    End If
    
    Do
        StartPos = InStr(TextToSearch, "**")
        If StartPos = 0 Then Exit Do
        
        ' Check if this is actually *** (handled by ProcessBoldItalic)
        If Mid(TextToSearch, StartPos, 3) = "***" Then
            ' Skip this occurrence - it's bold+italic, not just bold
            TextToSearch = Left(TextToSearch, StartPos) & "XX" & Mid(TextToSearch, StartPos + 2)
            GoTo NextBoldCheck
        End If
        
        EndPos = InStr(StartPos + 2, TextToSearch, "**")
        If EndPos = 0 Then Exit Do
        
        ' Check if ending is actually ***
        If Mid(TextToSearch, EndPos, 3) = "***" Then
            ' Skip - this is bold+italic
            TextToSearch = Left(TextToSearch, EndPos) & "XX" & Mid(TextToSearch, EndPos + 2)
            GoTo NextBoldCheck
        End If
        
        ' Apply bold - but only to the actual paragraph range, not including vbCr
        Set rng = Para.Range
        If Right(rng.Text, 1) = vbCr Then
            rng.End = rng.End - 1  ' Don't include paragraph mark
        End If
        
        rng.Start = rng.Start + StartPos - 1
        rng.End = rng.Start + (EndPos - StartPos) + 2
        
        ' Remove markers and format
        rng.Text = Mid(rng.Text, 3, Len(rng.Text) - 4)
        rng.Font.Bold = True
        
        ' Refresh our search text
        Set rng = Para.Range
        TextToSearch = rng.Text
        If Right(TextToSearch, 1) = vbCr Then
            TextToSearch = Left(TextToSearch, Len(TextToSearch) - 1)
        End If
        
NextBoldCheck:
    Loop
End Sub

Sub ProcessItalic(Para As Paragraph)
    ' Process *text* or _text_ for italic (but not ** or *** which are handled elsewhere)
    ProcessItalicMarker Para, "*"
    ProcessItalicMarker Para, "_"
End Sub

Sub ProcessItalicMarker(Para As Paragraph, Marker As String)
    Dim rng As Range
    Dim StartPos As Long
    Dim EndPos As Long
    Dim TextToSearch As String
    
    Set rng = Para.Range
    
    ' Don't include the paragraph mark in our search
    TextToSearch = rng.Text
    If Right(TextToSearch, 1) = vbCr Then
        TextToSearch = Left(TextToSearch, Len(TextToSearch) - 1)
    End If
    
    Do
        StartPos = InStr(TextToSearch, Marker)
        If StartPos = 0 Then Exit Do
        
        ' Make sure it's not part of ** or *** or ___
        If StartPos > 1 Then
            If Mid(TextToSearch, StartPos - 1, 1) = Marker Then
                TextToSearch = Left(TextToSearch, StartPos) & Mid(TextToSearch, StartPos + 2)
                GoTo NextItalicCheck
            End If
        End If
        
        ' Check if it's actually ** or ***
        If Mid(TextToSearch, StartPos, 2) = String(2, Marker) Then
            TextToSearch = Left(TextToSearch, StartPos) & Mid(TextToSearch, StartPos + 2)
            GoTo NextItalicCheck
        End If
        
        EndPos = InStr(StartPos + 1, TextToSearch, Marker)
        If EndPos = 0 Then Exit Do
        
        ' Make sure ending marker is not part of double
        If EndPos < Len(TextToSearch) Then
            If Mid(TextToSearch, EndPos + 1, 1) = Marker Then
                TextToSearch = Left(TextToSearch, EndPos) & Mid(TextToSearch, EndPos + 2)
                GoTo NextItalicCheck
            End If
        End If
        
        ' Apply italic - but only to the actual paragraph range, not including vbCr
        Set rng = Para.Range
        If Right(rng.Text, 1) = vbCr Then
            rng.End = rng.End - 1  ' Don't include paragraph mark
        End If
        
        rng.Start = rng.Start + StartPos - 1
        rng.End = rng.Start + (EndPos - StartPos) + 1
        
        ' Remove markers and format
        rng.Text = Mid(rng.Text, 2, Len(rng.Text) - 2)
        rng.Font.Italic = True
        
        ' Reset for next iteration
        Set rng = Para.Range
        TextToSearch = rng.Text
        If Right(TextToSearch, 1) = vbCr Then
            TextToSearch = Left(TextToSearch, Len(TextToSearch) - 1)
        End If
        
NextItalicCheck:
    Loop
End Sub

Sub ProcessInlineCode(Para As Paragraph)
    ' Process `code` for inline code character style
    Dim rng As Range
    Dim StartPos As Long
    Dim EndPos As Long
    Dim TextToSearch As String
    
    Set rng = Para.Range
    
    ' Don't include the paragraph mark in our search
    TextToSearch = rng.Text
    If Right(TextToSearch, 1) = vbCr Then
        TextToSearch = Left(TextToSearch, Len(TextToSearch) - 1)
    End If
    
    Do
        StartPos = InStr(TextToSearch, "`")
        If StartPos = 0 Then Exit Do
        
        EndPos = InStr(StartPos + 1, TextToSearch, "`")
        If EndPos = 0 Then Exit Do
        
        ' Apply Code style - but only to the actual paragraph range, not including vbCr
        Set rng = Para.Range
        If Right(rng.Text, 1) = vbCr Then
            rng.End = rng.End - 1  ' Don't include paragraph mark
        End If
        
        rng.Start = rng.Start + StartPos - 1
        rng.End = rng.Start + (EndPos - StartPos) + 2
        
        ' Remove markers and apply style
        rng.Text = Mid(rng.Text, 2, Len(rng.Text) - 2)
        
        On Error Resume Next
        rng.Style = "Code"
        On Error GoTo 0
        
        ' Refresh our search text
        Set rng = Para.Range
        TextToSearch = rng.Text
        If Right(TextToSearch, 1) = vbCr Then
            TextToSearch = Left(TextToSearch, Len(TextToSearch) - 1)
        End If
    Loop
End Sub

Sub ProcessLinks(Para As Paragraph)
    ' Process [label](url) for hyperlinks
    Dim rng As Range
    Dim StartPos As Long
    Dim MidPos As Long
    Dim EndPos As Long
    Dim LinkText As String
    Dim LinkUrl As String
    
    Set rng = Para.Range
    
    Do
        StartPos = InStr(rng.Text, "[")
        If StartPos = 0 Then Exit Do
        
        MidPos = InStr(StartPos, rng.Text, "](")
        If MidPos = 0 Then Exit Do
        
        EndPos = InStr(MidPos, rng.Text, ")")
        If EndPos = 0 Then Exit Do
        
        ' Extract link text and URL
        LinkText = Mid(rng.Text, StartPos + 1, MidPos - StartPos - 1)
        LinkUrl = Mid(rng.Text, MidPos + 2, EndPos - MidPos - 2)
        
        ' Create hyperlink
        Set rng = Para.Range
        rng.Start = rng.Start + StartPos - 1
        rng.End = rng.Start + (EndPos - StartPos) + 1
        
        rng.Text = LinkText
        ActiveDocument.Hyperlinks.Add Anchor:=rng, Address:=LinkUrl
        
        Set rng = Para.Range
    Loop
End Sub

Sub ProcessBareUrls(Para As Paragraph)
    ' Process bare URLs (http://, https://, www.)
    Dim rng As Range
    Dim Words As Variant
    Dim Word As Variant
    Dim I As Long
    
    Set rng = Para.Range
    
    ' Simple approach: look for http:// or https:// or www.
    For I = 1 To rng.Words.count
        If InStr(LCase(rng.Words(I).Text), "http://") = 1 Or _
           InStr(LCase(rng.Words(I).Text), "https://") = 1 Or _
           InStr(LCase(rng.Words(I).Text), "www.") = 1 Then
            
            ' Check if not already a hyperlink
            If rng.Words(I).Hyperlinks.count = 0 Then
                ActiveDocument.Hyperlinks.Add Anchor:=rng.Words(I), _
                                              Address:=Trim(rng.Words(I).Text)
            End If
        End If
    Next I
End Sub

Sub ProcessImages(Doc As Document)
    ' Process ![alt](url) for embedded images
    Dim Para As Paragraph
    Dim Text As String
    Dim StartPos As Long
    Dim MidPos As Long
    Dim EndPos As Long
    Dim AltText As String
    Dim ImageUrl As String
    Dim rng As Range
    Dim InlineShape As InlineShape
    Dim ParaCount As Long
    Dim TotalParas As Long
    
    TotalParas = Doc.Paragraphs.count
    ParaCount = 0
    
    On Error Resume Next
    
    For Each Para In Doc.Paragraphs
        ParaCount = ParaCount + 1
        
        ' Update status every 10 paragraphs
        If ParaCount Mod 10 = 0 Then
            Application.StatusBar = "Processing images: paragraph " & ParaCount & " of " & TotalParas
        End If
        
        ' Skip code block table cells (they have "Code" style applied)
        If Para.Range.Tables.count > 0 Then
            If Para.Style = "Code" Or Para.Range.Style = "Code" Then
                GoTo NextParaImage
            End If
        End If
        
        Set rng = Para.Range
        Text = rng.Text
        
        Do
            StartPos = InStr(Text, "![")
            If StartPos = 0 Then Exit Do
            
            MidPos = InStr(StartPos, Text, "](")
            If MidPos = 0 Then Exit Do
            
            EndPos = InStr(MidPos, Text, ")")
            If EndPos = 0 Then Exit Do
            
            ' Extract alt text and URL
            AltText = Mid(Text, StartPos + 2, MidPos - StartPos - 2)
            ImageUrl = Mid(Text, MidPos + 2, EndPos - MidPos - 2)
            
            Application.StatusBar = "Embedding image from: " & Left(ImageUrl, 50) & "..."
            
            ' Insert image
            Set rng = Para.Range
            rng.Start = rng.Start + StartPos - 1
            rng.End = rng.Start + (EndPos - StartPos) + 1
            
            Set InlineShape = rng.InlineShapes.AddPicture(filename:=ImageUrl, _
                                                          LinkToFile:=False, _
                                                          SaveWithDocument:=True)
            If Not InlineShape Is Nothing Then
                InlineShape.AlternativeText = AltText
            End If
            
            Text = Para.Range.Text
        Loop
        
NextParaImage:
    Next Para
    
    On Error GoTo 0
End Sub

' ==============================================================================
' NOTE: This macro requires two helper subs to exist in your project:
'   - InsertMyBulletList() - Creates "JDM Bullets" list style
'   - InsertJDMnumList()   - Creates "JDM 1.1)" list style
' ==============================================================================

