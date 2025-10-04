Attribute VB_Name = "Module1"
Option Explicit

' ==============================
' QUICK STYLE DIAGNOSTICS
' ==============================
' Dumps Paragraph, Table, and List styles to the Immediate window (Ctrl+G)

Public Sub DebugPrintStyles()
    Debug.Print String(60, "-")
    Debug.Print "STYLE DIAGNOSTICS @ " & Now
    On Error Resume Next
    Debug.Print "Template: ", ActiveDocument.AttachedTemplate.FullName
    On Error GoTo 0

    Debug.Print "Required:"
    Debug.Print "  Separator (Paragraph): ", ExistsLabel(StyleExistsQuick("Separator", wdStyleTypeParagraph))
    Debug.Print "  DW Array (Table):      ", ExistsLabel(StyleExistsQuick("DW Array", wdStyleTypeTable))
    Debug.Print "  JDM Bullet (Para):     ", ExistsLabel(StyleExistsQuick("JDM Bullet", wdStyleTypeParagraph))
    Debug.Print "  JDM 1.1) (Para):       ", ExistsLabel(StyleExistsQuick("JDM 1.1)", wdStyleTypeParagraph))

    Dim s As Style
    Debug.Print "Paragraph styles:"
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Then Debug.Print "  - ", s.NameLocal
    Next s

    Debug.Print "List styles:"
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeList Then Debug.Print "  - ", s.NameLocal
    Next s

    Debug.Print "Table styles:"
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeTable Then Debug.Print "  - ", s.NameLocal
    Next s

    Debug.Print String(60, "-")
End Sub

Private Function StyleExistsQuick(ByVal StyleName As String, ByVal styleType As WdStyleType) As Boolean
    On Error Resume Next
    Dim st As Style
    Set st = ActiveDocument.Styles(StyleName)
    If Not st Is Nothing Then
        StyleExistsQuick = (st.Type = styleType Or styleType = wdStyleTypeParagraph Or styleType = wdStyleTypeCharacter Or styleType = wdStyleTypeTable Or styleType = wdStyleTypeList)
    Else
        StyleExistsQuick = False
    End If
    On Error GoTo 0
End Function

Private Function ExistsLabel(ByVal b As Boolean) As String
    ExistsLabel = IIf(b, "FOUND", "MISSING")
End Function
