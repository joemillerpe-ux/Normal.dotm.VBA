Attribute VB_Name = "AI_WordToAI"
Sub AIxPrepSaveCloseAndCopyPath()
    Dim FilePath As String
    Dim LocalPath As String
    Dim SharePointPrefix As String
    Dim LocalPrefix As String

    ' Get the current file path
    FilePath = ActiveDocument.FullName

    ' Define the SharePoint and local path prefixes
    SharePointPrefix = "https://gibraltar1-my.sharepoint.com/personal/jmiller1_dsb_gibraltar1_com/Documents/Documents/"
    LocalPrefix = "C:\Users\jmiller1\OneDrive - Gibraltar Industries\Documents\"

    ' Apply the transformation
    LocalPath = Replace(FilePath, SharePointPrefix, LocalPrefix)
    LocalPath = Replace(LocalPath, "/", "\")

    ' Expand all headings
    ActiveDocument.ActiveWindow.View.ExpandAllHeadings

    ' Save current document
    ActiveDocument.Save

    ' Copy the transformed path using selection
    With Selection
        .EndKey Unit:=wdStory
        .TypeText vbCr & LocalPath
        .MoveLeft Unit:=wdCharacter, count:=Len(LocalPath), Extend:=wdExtend
        .Copy
        .TypeBackspace
    End With

    ' Notify the user
    MsgBox "File saved and path copied to clipboard." & vbCrLf & vbCrLf & "Please close MS-Word so you can reload the file to GPT." & vbCrLf & vbCrLf & LocalPath, vbInformation, "Save Complete"

    ' Close Word
    'Application.Quit SaveChanges:=wdSaveChanges
End Sub



