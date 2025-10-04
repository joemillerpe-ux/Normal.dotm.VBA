VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ===== CONTENT CONTROL CLIPBOARD COPY =====
' IMPORTANT: This code must be placed in ThisDocument module (not a regular module)
'
' To install:
' 1. Open VBA Editor (Alt+F11)
' 2. In Project Explorer, expand your document project
' 3. Double-click "ThisDocument"
' 4. Paste this entire code block there

Private Sub Document_ContentControlOnEnter(ByVal ContentControl As ContentControl)
'
' Document_ContentControlOnEnter Event Handler
' Automatically copies content control text to clipboard when clicked
' This event fires whenever user clicks into any content control
'
    CopyContentControlToClipboard ContentControl
End Sub

Private Sub CopyContentControlToClipboard(cc As ContentControl)
'
' CopyContentControlToClipboard Subroutine
' Copies content control text to clipboard and provides user feedback
' Handles empty content controls gracefully
'
    Dim textToCopy As String
    
    ' Get the text content and clean it
    textToCopy = Trim(cc.Range.Text)
    
    ' Skip if empty or just whitespace
    If Len(textToCopy) = 0 Then Exit Sub
    
    ' Copy to system clipboard using DataObject
    Dim dataObj As Object
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText textToCopy
    dataObj.PutInClipboard
    
    ' Provide user feedback via status bar
    Application.StatusBar = "Copied to clipboard: " & Left(textToCopy, 50) & IIf(Len(textToCopy) > 50, "...", "")
    
    ' Schedule status bar clearing after 2 seconds
    Application.OnTime Now + TimeValue("00:00:02"), "ClearContentControlStatusBar"
    
    ' Cleanup
    Set dataObj = Nothing
End Sub

Private Sub ClearContentControlStatusBar()
'
' ClearContentControlStatusBar Subroutine
' Helper function to clear the status bar message after delay
' Called automatically by Application.OnTime
'
    Application.StatusBar = ""
End Sub
