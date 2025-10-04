Attribute VB_Name = "Formatting"
Sub HighlightYellow()
  Selection.Range.HighlightColorIndex = wdYellow
End Sub
Sub HighlightGreen()
  Selection.Range.HighlightColorIndex = wdBrightGreen
End Sub
Sub HighlightRemoval()
  Selection.Range.HighlightColorIndex = wdNone
End Sub
Sub HighlightRemovalAll()
  Selection.WholeStory
  Selection.Range.HighlightColorIndex = wdNone
End Sub
Sub LastPositionCtrlAltZ()
  Application.GoBack
End Sub
Sub InsertMyBulletList()
'If list is not already defined, create it.  To redefine, use Sub DeleteListStyleByName.
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
    .NumberFormat = ChrW(61492)
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleBullet
    .NumberPosition = InchesToPoints(0)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(0.25)
    .TabPosition = wdUndefined
    .ResetOnHigher = 0
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = "Webdings"
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
    .NumberFormat = ChrW(61492)
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleBullet
    .NumberPosition = InchesToPoints(0.25)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(0.5)
    .TabPosition = wdUndefined
    .ResetOnHigher = 1
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = "Webdings"
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3)
    .NumberFormat = ChrW(61492)
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleBullet
    .NumberPosition = InchesToPoints(0.5)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(0.75)
    .TabPosition = wdUndefined
    .ResetOnHigher = 2
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = "Webdings"
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(4)
    .NumberFormat = ChrW(61492)
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleBullet
    .NumberPosition = InchesToPoints(0.75)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(1)
    .TabPosition = wdUndefined
    .ResetOnHigher = 3
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = "Webdings"
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(5)
    .NumberFormat = ChrW(61492)
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleBullet
    .NumberPosition = InchesToPoints(1)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(1.25)
    .TabPosition = wdUndefined
    .ResetOnHigher = 4
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = "Webdings"
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(6)
    .NumberFormat = ChrW(61492)
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleBullet
    .NumberPosition = InchesToPoints(1.25)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(1.5)
    .TabPosition = wdUndefined
    .ResetOnHigher = 5
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = "Webdings"
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(7)
    .NumberFormat = ChrW(61492)
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleBullet
    .NumberPosition = InchesToPoints(1.5)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(1.75)
    .TabPosition = wdUndefined
    .ResetOnHigher = 6
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = "Webdings"
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(8)
    .NumberFormat = ChrW(61492)
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleBullet
    .NumberPosition = InchesToPoints(1.75)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(2)
    .TabPosition = wdUndefined
    .ResetOnHigher = 7
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = "Webdings"
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(9)
    .NumberFormat = ChrW(61492)
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleBullet
    .NumberPosition = InchesToPoints(2)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(2.25)
    .TabPosition = wdUndefined
    .ResetOnHigher = 8
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = "Webdings"
    End With
    .LinkedStyle = ""
  End With
  ListGalleries(wdOutlineNumberGallery).ListTemplates(1).name = _
    "JDM Bullets"
  Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
    ContinuePreviousList:=True, ApplyTo:=wdListApplyToWholeList, _
    DefaultListBehavior:=wdWord10ListBehavior
End Sub
Sub InsertJDMnumList()
'If list is not already defined, create it.  To redefine, use Sub DeleteListStyleByName.
Dim listStyleName As String
listStyleName = "ListJDM1.1"

On Error Resume Next ' Enable error handling
    Selection.Style = ActiveDocument.Styles(listStyleName)
    
    If Err.Number <> 0 Then
'        MsgBox "The specified list style does not exist. Adding now...", vbExclamation
        Err.Clear
    Else: Exit Sub
    End If
    On Error GoTo 0 ' Disable error handling

'code to create same list style
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(1)
    .NumberFormat = "%1)"
    .NumberStyle = wdListNumberStyleArabic
    .TrailingCharacter = wdTrailingSpace
    .NumberPosition = InchesToPoints(0)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(0.25)
    .TabPosition = wdUndefined
    .ResetOnHigher = 0
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = ""
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(2)
    .NumberFormat = "%1.%2)"
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = InchesToPoints(0.25)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(0.5)
    .TabPosition = wdUndefined
    .ResetOnHigher = 1
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = ""
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(3)
    .NumberFormat = "%1.%2.%3)"
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = InchesToPoints(0.5)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(0.75)
    .TabPosition = wdUndefined
    .ResetOnHigher = 2
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = ""
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(4)
    .NumberFormat = "%1.%2.%3.%4)"
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = InchesToPoints(0.75)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(1)
    .TabPosition = wdUndefined
    .ResetOnHigher = 3
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = ""
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(5)
    .NumberFormat = "%1.%2.%3.%4.%5)"
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = InchesToPoints(1)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(1.25)
    .TabPosition = wdUndefined
    .ResetOnHigher = 4
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = ""
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(6)
    .NumberFormat = "%1.%2.%3.%4.%5.%6)"
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = InchesToPoints(1.25)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(1.5)
    .TabPosition = wdUndefined
    .ResetOnHigher = 5
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = ""
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(7)
    .NumberFormat = "%1.%2.%3.%4.%5.%6.%7)"
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = InchesToPoints(1.5)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(1.75)
    .TabPosition = wdUndefined
    .ResetOnHigher = 6
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = ""
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(8)
    .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8)"
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = InchesToPoints(1.75)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(2)
    .TabPosition = wdUndefined
    .ResetOnHigher = 7
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = ""
    End With
    .LinkedStyle = ""
  End With
  With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(9)
    .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9)"
    .TrailingCharacter = wdTrailingSpace
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = InchesToPoints(2)
    .Alignment = wdListLevelAlignLeft
    .TrailingCharacter = wdTrailingSpace
    .TextPosition = InchesToPoints(2.25)
    .TabPosition = wdUndefined
    .ResetOnHigher = 8
    .StartAt = 1
    With .Font
      .Bold = wdUndefined
      .Italic = wdUndefined
      .StrikeThrough = wdUndefined
      .Subscript = wdUndefined
      .Superscript = wdUndefined
      .Shadow = wdUndefined
      .Outline = wdUndefined
      .Emboss = wdUndefined
      .Engrave = wdUndefined
      .AllCaps = wdUndefined
      .Hidden = wdUndefined
      .Underline = wdUndefined
      .Color = wdUndefined
      .Size = wdUndefined
      .Animation = wdUndefined
      .DoubleStrikeThrough = wdUndefined
      .name = ""
    End With
    .LinkedStyle = ""
  End With
  ListGalleries(wdOutlineNumberGallery).ListTemplates(2).name = ""
  Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
    ListGalleries(wdOutlineNumberGallery).ListTemplates(2), _
    ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, _
    DefaultListBehavior:=wdWord10ListBehavior
    
End Sub
Sub DeleteListStyleByName()
    Dim StyleName As String
    Dim s As Style

    StyleName = "ListJDM1.1"

    On Error Resume Next
    Set s = ActiveDocument.Styles(StyleName)
    If Not s Is Nothing Then
        If s.Type = wdStyleTypeList Then
            ActiveDocument.Styles(StyleName).Delete
            MsgBox "List style '" & StyleName & "' deleted."
        Else
            MsgBox "'" & StyleName & "' is not a list style."
        End If
    Else
        MsgBox "List style '" & StyleName & "' not found."
    End If
    On Error GoTo 0
End Sub

