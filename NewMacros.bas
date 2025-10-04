Attribute VB_Name = "NewMacros"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
  Selection.ConvertToTable Separator:=wdSeparateByParagraphs, NumColumns:=1, _
     NumRows:=1, AutoFitBehavior:=wdAutoFitFixed
  With Selection.Tables(1)
    .Style = "Table Grid"
    .ApplyStyleHeadingRows = True
    .ApplyStyleLastRow = False
    .ApplyStyleFirstColumn = True
    .ApplyStyleLastColumn = False
  End With
  Selection.Tables(1).rows.WrapAroundText = True
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=715.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=710.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=706.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=701.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=697.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=692.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=688.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=683.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=679.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=674.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=661.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=638.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=616.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=593.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=566.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=512.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=458.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=386.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=346.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=314.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=287.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=260.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=220.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=193.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=161.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=143.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=130.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=116.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=98.75, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=85.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=67.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=58.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=58.25, RulerStyle:= _
    wdAdjustFirstColumn
  Selection.Delete Unit:=wdCharacter, count:=1
  Selection.Delete Unit:=wdCharacter, count:=1
  Selection.Delete Unit:=wdCharacter, count:=1
  Selection.Delete Unit:=wdCharacter, count:=1
  Selection.Delete Unit:=wdCharacter, count:=1
  Selection.Delete Unit:=wdCharacter, count:=1
  Selection.Delete Unit:=wdCharacter, count:=1
  With Selection.Tables(1).rows
    .WrapAroundText = True
    .HorizontalPosition = InchesToPoints(0.1)
    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    .DistanceLeft = InchesToPoints(0.13)
    .DistanceRight = InchesToPoints(0.13)
    .VerticalPosition = InchesToPoints(0)
    .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
    .DistanceTop = InchesToPoints(0)
    .DistanceBottom = InchesToPoints(0)
    .AllowOverlap = False
  End With
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
'
  Selection.TypeText Text:="TEMP "
  Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
    PreserveFormatting:=False
  Selection.Fields.Update
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro3"
'
' Macro3 Macro
'
'
  Selection.Range.ContentControls.Add (wdContentControlText)
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro4"
'
' Macro4 Macro
'
'
  Selection.Style = ActiveDocument.Styles("Normal")
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro5"
'
' Macro5 Macro
'
'
  Selection.InlineShapes.AddHorizontalLineStandard
End Sub
