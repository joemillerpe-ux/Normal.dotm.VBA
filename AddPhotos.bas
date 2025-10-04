Attribute VB_Name = "AddPhotos"
Sub AddPics()

Application.ScreenUpdating = False

Dim I As Long, j As Long, C As Long, r As Long, NumCols As Long, iShp As InlineShape

Dim oTbl As Table, TblWdth As Single, StrTxt As String, RwHght As Single, ColWdth As Single

On Error GoTo ErrExit

NumCols = 1

RwHght = ((11 - 0.75 - 0.2 - 0.2 - 0.25) * 2.54) / 2 '9.525 'cm

On Error GoTo 0

'Select and insert the Pics

With Application.FileDialog(msoFileDialogFilePicker)

  .title = "Select image files and click OK"

  .Filters.Add "Images", "*.gif; *.jpg; *.jpeg; *.bmp; *.tif; *.png"

  .FilterIndex = 2

  If .Show = -1 Then

    'Create a paragraph Style with 0 space before/after & centre-aligned

    On Error Resume Next

    With ActiveDocument

      .Styles.Add name:="TblPic", Type:=wdStyleTypeParagraph

      On Error GoTo 0

      With .Styles("TblPic").ParagraphFormat

        .Alignment = wdAlignParagraphCenter

        .KeepWithNext = True

        .SpaceAfter = 0

        .SpaceBefore = 0

      End With

    End With

    'Add a 2-row by NumCols-column table to take the images

    Set oTbl = Selection.Tables.Add(Range:=Selection.Range, NumRows:=2, NumColumns:=NumCols)

    With ActiveDocument.PageSetup

      TblWdth = .PageWidth - .LeftMargin - .RightMargin - .Gutter

      ColWdth = TblWdth / NumCols

    End With

    With oTbl

      .AutoFitBehavior (wdAutoFitFixed)

      .Columns.Width = ColWdth

    End With

    CaptionLabels.Add name:="Picture"

     For I = 1 To .SelectedItems.count Step NumCols

      r = ((I - 1) / NumCols + 1) * 2 - 1

      'Format the rows

      Call FormatRows(oTbl, r, RwHght)

      For C = 1 To NumCols

        j = j + 1

        'Insert the Picture

        Set iShp = ActiveDocument.InlineShapes.AddPicture( _
          filename:=.SelectedItems(j), LinkToFile:=False, _
          SaveWithDocument:=True, Range:=oTbl.Cell(r, C).Range)

        With iShp

          .LockAspectRatio = True

          If (.Width < ColWdth) And (.Height < RwHght) Then

            .Width = ColWdth

            If .Height > RwHght Then .Height = RwHght

          End If

        End With

        'Get the Image name for the Caption

        StrTxt = Split(.SelectedItems(j), "\")(UBound(Split(.SelectedItems(j), "\")))

        StrTxt = ": " & Split(StrTxt, ".")(0)

        'Insert the Caption on the row below the picture

        With oTbl.Cell(r + 1, C).Range

          .InsertBefore vbCr

          .Characters.First.InsertCaption _
          label:="Picture", title:=StrTxt, _
          Position:=wdCaptionPositionBelow, ExcludeLabel:=True

          .Characters.First = vbNullString

          .Characters.Last.Previous = vbNullString

        

        End With

        'Exit when we're done

        If j = .SelectedItems.count Then Exit For

      Next

      'Add extra rows as needed

      If j < .SelectedItems.count Then

        oTbl.rows.Add

        oTbl.rows.Add

      End If

    Next

  Else

   End If

End With

ErrExit:

Application.ScreenUpdating = True

End Sub


Sub FormatRows(oTbl As Table, x As Long, Hght As Single)

With oTbl

  With .rows(x)

    .Height = CentimetersToPoints(Hght)

    .HeightRule = wdRowHeightExactly

    .Range.Style = "TblPic"

    .Cells.VerticalAlignment = wdCellAlignVerticalCenter

  End With

  With .rows(x + 1)

    .Height = CentimetersToPoints(0.5)

    .HeightRule = wdRowHeightExactly

    .Range.Style = "Caption"

  End With

End With

End Sub

