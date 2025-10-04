Attribute VB_Name = "Module3"
' ==============================
' ApplyFormattingStandards — 5017 DIAGNOSTICS PACK
' ==============================
' Goal: pinpoint the exact phase/line raising "Run-time error 5017: Application-defined or object-defined error"
' Strategy:
'   1) Add a tiny logger that writes to a Desktop log file (per your prefs).
'   2) Instrument the high-risk sections (EnsureRequiredStyles, list-style builders, table-style creation,
'      Heading1 lettering) with try/catch + Log.
'   3) Re-run, then share the last ~80 lines of the log so we can fix precisely.
'
' This pack contains:
'   • Module: JDM_Logger (new)
'   • Patches to: ApplyFormattingStandards, EnsureRequiredStyles, AutoCreateStyles,
'                 NumberHeading1LettersWithSectionRestarts
'
' ==============================
' Module: JDM_Logger (NEW)
' ==============================
Option Explicit

Private gLogPath As String

Public Sub StartLog(Optional ByVal title As String = "ApplyFormattingStandards")
    On Error Resume Next
    gLogPath = GetDesktopPath() & "\" & title & "_" & format(Now, "yyyymmdd_HHNNSS") & ".log"
    Dim ff As Integer: ff = FreeFile
    Open gLogPath For Output As #ff
    Print #ff, Now & " — LOG START: " & title
    Close #ff
    On Error GoTo 0
End Sub

Public Sub LogMsg(ByVal msg As String)
    On Error Resume Next
    If Len(gLogPath) = 0 Then StartLog "ApplyFormattingStandards"
    Dim ff As Integer: ff = FreeFile
    Open gLogPath For Append As #ff
    Print #ff, Now & " — " & msg
    Close #ff
    On Error GoTo 0
End Sub

Public Sub LogErr(ByVal ctx As String)
    On Error Resume Next
    LogMsg "ERR in " & ctx & ": " & Err.Number & " — " & Err.Description
    On Error GoTo 0
End Sub

Public Sub EndLog()
    On Error Resume Next
    Dim ff As Integer: ff = FreeFile
    Open gLogPath For Append As #ff
    Print #ff, Now & " — LOG END"
    Close #ff
    On Error GoTo 0
End Sub

Private Function GetDesktopPath() As String
    On Error Resume Next
    Dim p As String
    p = Environ$("USERPROFILE") & "\Desktop"
    If Len(Dir$(p, vbDirectory)) = 0 Then
        ' Fallback
        p = CurDir$
    End If
    GetDesktopPath = p
    On Error GoTo 0
End Function

' =====================================
' Patches in existing module (DIFFS)
' =====================================

' ---------- BEFORE ----------
' Public Sub ApplyFormattingStandards()
'     Dim undoRec As UndoRecord
'     On Error GoTo Fail
'     Set undoRec = Application.UndoRecord
'     undoRec.StartCustomRecord "ApplyFormattingStandards"
'     Dim doc As Document
'     Set doc = ActiveDocument
'     If Not EnsureRequiredStyles(doc) Then GoTo Cleanup
'     ' ...
' Cleanup:
'     If Not undoRec Is Nothing Then undoRec.EndCustomRecord
'     Exit Sub
' Fail:
'     If Not undoRec Is Nothing Then undoRec.EndCustomRecord
'     MsgBox "ApplyFormattingStandards error: " & Err.Number & " - " & Err.Description, vbExclamation
' End Sub

' ---------- AFTER (INSTRUMENTED) ----------
' Public Sub ApplyFormattingStandards()
'     Dim undoRec As UndoRecord
'     On Error GoTo Fail
'     StartLog "ApplyFormattingStandards"  ' NEW
'     LogMsg "Begin ApplyFormattingStandards"
'     Set undoRec = Application.UndoRecord
'     undoRec.StartCustomRecord "ApplyFormattingStandards"
'     Dim doc As Document
'     Set doc = ActiveDocument
'     LogMsg "ActiveDocument obtained"
'
'     LogMsg "EnsureRequiredStyles ?"
'     If Not EnsureRequiredStyles(doc) Then GoTo Cleanup
'     LogMsg "EnsureRequiredStyles ?"
'
'     LogMsg "NormalizeNewlines ?"
'     NormalizeNewlines doc
'     LogMsg "NormalizeNewlines ?"
'
'     LogMsg "ConvertFencedCodeBlocks ?"
'     ConvertFencedCodeBlocks doc
'     LogMsg "ConvertMarkdownImages ?"
'     ConvertMarkdownImages doc
'     LogMsg "ConvertMarkdownLinks ?"
'     ConvertMarkdownLinks doc
'     LogMsg "ConvertHorizontalRules ?"
'     ConvertHorizontalRules doc
'     LogMsg "ConvertPipeTables ?"
'     ConvertPipeTables doc
'     LogMsg "MapHeadings ?"
'     MapHeadings doc
'     LogMsg "ConvertBlockquotes ?"
'     ConvertBlockquotes doc
'     LogMsg "ApplyInlineFormatting ?"
'     ApplyInlineFormatting doc
'     LogMsg "ConvertLists ?"
'     ConvertLists doc
'     LogMsg "NumberHeading1LettersWithSectionRestarts ?"
'     On Error Resume Next
'     NumberHeading1LettersWithSectionRestarts doc
'     If Err.Number <> 0 Then LogErr "NumberHeading1LettersWithSectionRestarts": Err.Clear
'     On Error GoTo Fail
'     LogMsg "PromptForAdditionalHeading1Restarts ?"
'     PromptForAdditionalHeading1Restarts doc
'
' Cleanup:
'     If Not undoRec Is Nothing Then undoRec.EndCustomRecord
'     LogMsg "ApplyFormattingStandards cleanup"
'     EndLog
'     Exit Sub
' Fail:
'     If Not undoRec Is Nothing Then undoRec.EndCustomRecord
'     LogErr "ApplyFormattingStandards (Fail)"
'     EndLog
'     MsgBox "ApplyFormattingStandards error: " & Err.Number & " - " & Err.Description, vbExclamation
' End Sub

' ---------- BEFORE ----------
' Private Function EnsureRequiredStyles(ByVal doc As Document) As Boolean
'     ' ...
'     requiredListStyles = Array("JDM Bullets", "JDM 1.1)")
'     ' If Not StyleExists(...)
'     '   InsertMyBulletList / InsertJDMnumList
'     ' ...
' End Function

' ---------- AFTER (INSTRUMENTED + SAFEGUARDS) ----------
' Private Function EnsureRequiredStyles(ByVal doc As Document) As Boolean
'     Dim missing As Collection
'     Set missing = New Collection
'
'     Dim requiredParaStyles As Variant
'     requiredParaStyles = Array("Title", "Normal", "Quote", "Separator", "Code")
'     Dim requiredCharStyles As Variant
'     requiredCharStyles = Array("Code")
'     Dim requiredTableStyles As Variant
'     requiredTableStyles = Array("DW Array")
'     Dim requiredListStyles As Variant
'     requiredListStyles = Array("JDM Bullets", "JDM 1.1)")
'
'     LogMsg "EnsureRequiredStyles: pre-create checks"
'
'     ' Try to create the list styles via your builders; log any errors
'     If Not StyleExists(doc, "JDM Bullets", wdStyleTypeParagraph) Then
'         On Error Resume Next
'         LogMsg "Calling InsertMyBulletList…"
'         InsertMyBulletList
'         If Err.Number <> 0 Then LogErr "InsertMyBulletList": Err.Clear
'         On Error GoTo 0
'     End If
'     If Not StyleExists(doc, "JDM 1.1)", wdStyleTypeParagraph) Then
'         On Error Resume Next
'         LogMsg "Calling InsertJDMnumList…"
'         InsertJDMnumList
'         If Err.Number <> 0 Then LogErr "InsertJDMnumList": Err.Clear
'         On Error GoTo 0
'     End If
'
'     Dim name As Variant
'     For Each name In requiredParaStyles
'         If Not StyleExists(doc, name, wdStyleTypeParagraph) Then missing.Add "Paragraph: " & CStr(name)
'     Next name
'
'     For Each name In requiredCharStyles
'         If Not StyleExists(doc, name, wdStyleTypeCharacter) Then missing.Add "Character: " & CStr(name)
'     Next name
'
'     Dim i As Long
'     For i = LBound(requiredTableStyles) To UBound(requiredTableStyles)
'         If Not TableStyleExists(doc, CStr(requiredTableStyles(i))) Then missing.Add "Table: " & CStr(requiredTableStyles(i))
'     Next i
'
'     For Each name In requiredListStyles
'         If Not StyleExists(doc, name, wdStyleTypeParagraph) Then missing.Add "Paragraph (List): " & CStr(name)
'     Next name
'
'     LogMsg "EnsureRequiredStyles: missing count=" & CStr(missing.Count)
'
'     If missing.Count = 0 Then EnsureRequiredStyles = True: Exit Function
'
'     Dim listMissing As String, itm As Variant
'     listMissing = ""
'     For Each itm In missing
'         listMissing = listMissing & "• " & CStr(itm) & vbCrLf
'     Next itm
'
'     Dim resp As VbMsgBoxResult
'     resp = MsgBox("These required styles are missing:" & vbCrLf & vbCrLf & listMissing & vbCrLf & _
'                   "Create them now (Yes), continue but SKIP features needing them (No), or CANCEL altogether?", _
'                   vbYesNoCancel + vbInformation, "Missing Styles")
'     LogMsg "EnsureRequiredStyles: user resp=" & CStr(resp)
'     If resp = vbCancel Then
'         EnsureRequiredStyles = False
'         Exit Function
'     ElseIf resp = vbYes Then
'         On Error Resume Next
'         AutoCreateStyles doc, requiredParaStyles, requiredCharStyles, requiredTableStyles, requiredListStyles
'         If Err.Number <> 0 Then LogErr "AutoCreateStyles": Err.Clear
'         On Error GoTo 0
'     Else
'         ' Skip
'     End If
'
'     EnsureRequiredStyles = True
' End Function

' ---------- BEFORE ----------
' Private Sub AutoCreateStyles(...)
'     ' Creates styles; unguarded TableStyles.Condition(wdFirstRow)
' End Sub

' ---------- AFTER (HARDENED) ----------
' Private Sub AutoCreateStyles(ByVal doc As Document, _
'                              ByVal paraNames As Variant, _
'                              ByVal charNames As Variant, _
'                              ByVal tableNames As Variant, _
'                              ByVal listNames As Variant)
'     Dim nm As Variant
'     On Error Resume Next
'
'     For Each nm In paraNames
'         If Not StyleExists(doc, CStr(nm), wdStyleTypeParagraph) Then
'             LogMsg "Creating paragraph style: " & CStr(nm)
'             doc.Styles.Add Name:=CStr(nm), Type:=wdStyleTypeParagraph
'             If Err.Number <> 0 Then LogErr "Styles.Add para " & CStr(nm): Err.Clear
'             If CStr(nm) = "Separator" Then
'                 With doc.Styles("Separator").ParagraphFormat
'                     .SpaceBefore = 0
'                     .SpaceAfter = 6
'                     .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
'                     .Borders(wdBorderBottom).Color = wdColorGray40
'                 End With
'             ElseIf CStr(nm) = "Code" Then
'                 With doc.Styles("Code")
'                     .NoSpaceBetweenParagraphsOfSameStyle = True
'                     .Font.Name = "Consolas"
'                     .Font.Size = 10
'                     .ParagraphFormat.SpaceBefore = 0
'                     .ParagraphFormat.SpaceAfter = 0
'                 End With
'             End If
'         End If
'     Next nm
'
'     For Each nm In charNames
'         If Not StyleExists(doc, CStr(nm), wdStyleTypeCharacter) Then
'             LogMsg "Creating character style: " & CStr(nm)
'             doc.Styles.Add Name:=CStr(nm), Type:=wdStyleTypeCharacter
'             If Err.Number <> 0 Then LogErr "Styles.Add char " & CStr(nm): Err.Clear
'             If CStr(nm) = "Code" Then
'                 With doc.Styles("Code")
'                     .Font.Name = "Consolas"
'                     .Font.Size = 10
'                 End With
'             End If
'         End If
'     Next nm
'
'     Dim tn As Variant
'     For Each tn In tableNames
'         If Not TableStyleExists(doc, CStr(tn)) Then
'             LogMsg "Creating table style: " & CStr(tn)
'             doc.TableStyles.Add Name:=CStr(tn)
'             If Err.Number <> 0 Then LogErr "TableStyles.Add " & CStr(tn): Err.Clear
'             On Error Resume Next
'             With doc.TableStyles(CStr(tn))
'                 .AllowPageBreaks = True
'                 .Condition(wdFirstRow).Shading.BackgroundPatternColor = wdColorAutomatic
'                 .Condition(wdFirstRow).Font.Bold = True
'             End With
'             If Err.Number <> 0 Then LogErr "TableStyle.Condition " & CStr(tn): Err.Clear
'             On Error GoTo 0
'         End If
'     Next tn
'
'     For Each nm In listNames
'         If Not StyleExists(doc, CStr(nm), wdStyleTypeParagraph) Then
'             LogMsg "Creating list-paragraph style shell: " & CStr(nm)
'             doc.Styles.Add Name:=CStr(nm), Type:=wdStyleTypeParagraph
'             If Err.Number <> 0 Then LogErr "Styles.Add list para " & CStr(nm): Err.Clear
'         End If
'     Next nm
'
'     On Error GoTo 0
' End Sub

' ---------- HARDEN HEADING 1 LETTERING ----------
' Private Sub NumberHeading1LettersWithSectionRestarts(ByVal doc As Document)
'     LogMsg "NumberHeading1LettersWithSectionRestarts: begin"
'     On Error GoTo HandleErr
'     Dim sec As Section
'     For Each sec In doc.Sections
'         Dim r As Range
'         Set r = sec.Range
'         Dim p As Paragraph
'         Dim firstH1Found As Boolean: firstH1Found = False
'         For Each p In r.Paragraphs
'             If p.Range.Style Like "Heading 1" Then
'                 With p.Range.ListFormat
'                     .RemoveNumbers NumberType:=wdNumberParagraph
'                     .ApplyListTemplateWithLevel ListTemplate:=Nothing, ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList
'                     .ListTemplate.ListLevels(1).NumberFormat = "%1."
'                     .ListTemplate.ListLevels(1).NumberStyle = wdListNumberStyleUppercaseLetter
'                     .ListTemplate.ListLevels(1).ResetOnHigher = True
'                     .ListTemplate.ListLevels(1).LinkedStyle = "Heading 1"
'                 End With
'                 If Not firstH1Found Then
'                     p.Range.ListFormat.ApplyListTemplateWithLevel p.Range.ListFormat.ListTemplate, _
'                         ContinuePreviousList:=False, ApplyTo:=wdListApplyToThisPointForward, DefaultListBehavior:=wdWord10ListBehavior
'                     firstH1Found = True
'                 Else
'                     p.Range.ListFormat.ApplyListTemplateWithLevel p.Range.ListFormat.ListTemplate, _
'                         ContinuePreviousList:=True, ApplyTo:=wdListApplyToThisPointForward, DefaultListBehavior:=wdWord10ListBehavior
'                 End If
'             End If
'         Next p
'     Next sec
'     LogMsg "NumberHeading1LettersWithSectionRestarts: end"
'     Exit Sub
' HandleErr:
'     LogErr "NumberHeading1LettersWithSectionRestarts (core)"
' End Sub


