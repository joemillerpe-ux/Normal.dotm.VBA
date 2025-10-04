Attribute VB_Name = "Module2"
Sub ConvertAllFieldsToText()
    '
    ' ConvertAllFieldsToText Macro
    ' Converts all fields in the active document to plain text
    '
    
    Dim Doc As Document
    Dim fld As field
    Dim I As Integer
    
    Set Doc = ActiveDocument
    
    ' Check if there are any fields in the document
    If Doc.Fields.count = 0 Then
        MsgBox "No fields found in the document.", vbInformation
        Exit Sub
    End If
    
    ' Store the original count for user feedback
    Dim originalCount As Integer
    originalCount = Doc.Fields.count
    
    ' Start undo record for entire operation
    Doc.UndoRecord.StartCustomRecord "Convert All Fields to Text"
    
    ' Loop through all fields from last to first (to avoid index issues)
    For I = Doc.Fields.count To 1 Step -1
        Set fld = Doc.Fields(I)
        
        ' Update the field before converting (optional)
        fld.Update
        
        ' Convert field to text (unlink the field)
        fld.Unlink
    Next I
    
    ' End undo record
    Doc.UndoRecord.EndCustomRecord
    
    ' Provide user feedback
    MsgBox originalCount & " fields have been converted to text." & vbCrLf & _
           "Use Ctrl+Z to undo if needed.", vbInformation
    
End Sub

Sub ConvertSelectedFieldsToText()
    '
    ' ConvertSelectedFieldsToText Macro
    ' Converts only fields in the selected text to plain text
    '
    
    Dim sel As Selection
    Dim fld As field
    Dim I As Integer
    Dim Doc As Document
    
    Set sel = Selection
    Set Doc = ActiveDocument
    
    ' Check if text is selected
    If sel.Type = wdSelectionIP Then
        MsgBox "Please select text containing fields first.", vbExclamation
        Exit Sub
    End If
    
    ' Check if there are any fields in the selection
    If sel.Fields.count = 0 Then
        MsgBox "No fields found in the selected text.", vbInformation
        Exit Sub
    End If
    
    ' Store the original count for user feedback
    Dim originalCount As Integer
    originalCount = sel.Fields.count
    
    ' Start undo record for entire operation
    Doc.UndoRecord.StartCustomRecord "Convert Selected Fields to Text"
    
    ' Loop through all fields in selection from last to first
    For I = sel.Fields.count To 1 Step -1
        Set fld = sel.Fields(I)
        
        ' Update the field before converting (optional)
        fld.Update
        
        ' Convert field to text (unlink the field)
        fld.Unlink
    Next I
    
    ' End undo record
    Doc.UndoRecord.EndCustomRecord
    
    ' Provide user feedback
    MsgBox originalCount & " fields in selection have been converted to text." & vbCrLf & _
           "Use Ctrl+Z to undo if needed.", vbInformation
    
End Sub

Sub ConvertSpecificFieldTypeToText()
    '
    ' ConvertSpecificFieldTypeToText Macro
    ' Converts only specific types of fields to plain text
    ' Modify the field types as needed
    '
    
    Dim Doc As Document
    Dim fld As field
    Dim I As Integer
    Dim convertCount As Integer
    
    Set Doc = ActiveDocument
    convertCount = 0
    
    ' Check if there are any fields in the document
    If Doc.Fields.count = 0 Then
        MsgBox "No fields found in the document.", vbInformation
        Exit Sub
    End If
    
    ' Start undo record for entire operation
    Doc.UndoRecord.StartCustomRecord "Convert Specific Field Types to Text"
    
    ' Loop through all fields from last to first
    For I = Doc.Fields.count To 1 Step -1
        Set fld = Doc.Fields(I)
        
        ' Check for specific field types (modify as needed)
        Select Case fld.Type
            Case wdFieldDate, wdFieldTime, wdFieldPage, wdFieldNumPages
                ' Update the field before converting
                fld.Update
                ' Convert field to text
                fld.Unlink
                convertCount = convertCount + 1
                
            ' Add more field types as needed:
            ' Case wdFieldRef, wdFieldPageRef
            '     fld.Update
            '     fld.Unlink
            '     convertCount = convertCount + 1
        End Select
    Next I
    
    ' End undo record
    Doc.UndoRecord.EndCustomRecord
    
    ' Provide user feedback
    If convertCount > 0 Then
        MsgBox convertCount & " specific fields have been converted to text." & vbCrLf & _
               "Use Ctrl+Z to undo if needed.", vbInformation
    Else
        MsgBox "No matching field types found to convert.", vbInformation
    End If
    
End Sub

Sub ConvertFieldsInHeadersFooters()
    '
    ' ConvertFieldsInHeadersFooters Macro
    ' Converts fields in headers and footers to plain text
    '
    
    Dim Doc As Document
    Dim sec As Section
    Dim hdrFtr As HeaderFooter
    Dim fld As field
    Dim I As Integer, j As Integer
    Dim convertCount As Integer
    
    Set Doc = ActiveDocument
    convertCount = 0
    
    ' Start undo record for entire operation
    Doc.UndoRecord.StartCustomRecord "Convert Header/Footer Fields to Text"
    
    ' Loop through all sections
    For Each sec In Doc.Sections
        ' Loop through all headers and footers in each section
        For Each hdrFtr In sec.Headers
            If hdrFtr.Exists Then
                For j = hdrFtr.Range.Fields.count To 1 Step -1
                    Set fld = hdrFtr.Range.Fields(j)
                    fld.Update
                    fld.Unlink
                    convertCount = convertCount + 1
                Next j
            End If
        Next hdrFtr
        
        For Each hdrFtr In sec.Footers
            If hdrFtr.Exists Then
                For j = hdrFtr.Range.Fields.count To 1 Step -1
                    Set fld = hdrFtr.Range.Fields(j)
                    fld.Update
                    fld.Unlink
                    convertCount = convertCount + 1
                Next j
            End If
        Next hdrFtr
    Next sec
    
    ' End undo record
    Doc.UndoRecord.EndCustomRecord
    
    ' Provide user feedback
    If convertCount > 0 Then
        MsgBox convertCount & " fields in headers/footers have been converted to text." & vbCrLf & _
               "Use Ctrl+Z to undo if needed.", vbInformation
    Else
        MsgBox "No fields found in headers or footers.", vbInformation
    End If
    
End Sub


