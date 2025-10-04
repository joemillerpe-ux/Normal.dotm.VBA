Attribute VB_Name = "GitProject"
'Attribute VB_Name = "Project"
Option Explicit
''
' This module exports/imports VBA components between Normal.dotm and a file-system folder.
' Adapted so that:
'   - Source is Normal.dotm at:
'       C:\Users\jmiller1\AppData\Roaming\Microsoft\Templates\Normal.dotm
'   - Project root is:
'       C:\Users\jmiller1\AppData\Roaming\Microsoft\Templates\GithubNormalVBA
'   - Export/Import folder is:
'       <Project root>\src
'
' REQUIREMENTS:
'   - Word Options > Trust Center > Trust Center Settings > Macro Settings:
'       "Trust access to the VBA project object model" must be enabled.
'   - References:
'       * Microsoft Visual Basic for Applications Extensibility 5.3
'       * Microsoft Scripting Runtime
'
' NOTES:
'   - ExportComponentsToSourceFolder() clears the target src folder before exporting.
'   - Document-type components (vbext_ct_Document) cannot be programmatically removed.
'   - DangerouslyImportComponentsFromSourceFolder() imports from the same src folder.
''

' ======= PATHS (edit if needed) =======
Private Const NORMAL_DOTM_PATH As String = _
  "C:\Users\jmiller1\AppData\Roaming\Microsoft\Templates\Normal.dotm"

Private Const GITHUB_NORMAL_VBA_DIR As String = _
  "C:\Users\jmiller1\AppData\Roaming\Microsoft\Templates\GithubNormalVBA"

' ======= Root Directory of this "Project" (git repo root) =======
Public Property Get Dirname() As String
    ' If you prefer env var style, you could replace the constant with:
    ' Dirname = Environ$("AppData") & "\Microsoft\Templates\GithubNormalVBA"
    Dirname = GITHUB_NORMAL_VBA_DIR
End Property

' ======= Directory where all source code will be stored (./src) =======
Public Property Get SourceDirectory() As String
    SourceDirectory = joinPaths(Dirname, "src")
End Property

' ======= VBComponents for THIS project (Normal.dotm) =======
Private Property Get thisProjectsVBComponents() As VBComponents
    On Error GoTo Fallback
    If LCase$(Application.normalTemplate.FullName) = LCase$(NORMAL_DOTM_PATH) Then
        Set thisProjectsVBComponents = Application.normalTemplate.VBProject.VBComponents
        Exit Property
    End If
Fallback:
    Set thisProjectsVBComponents = Application.normalTemplate.VBProject.VBComponents
End Property

' ======= Helper: run scripts from the repo root =======
Public Function Bash(script As String, Optional keepCommandWindowOpen As Boolean = False) As Double
    Dim cmd As String
    cmd = "cmd.exe /S /" & IIf(keepCommandWindowOpen, "K", "C") & _
          " cd /d """ & Dirname & """ && " & script
    Bash = Shell(cmd, vbNormalFocus)
End Function

' ======= Initialize git in repo root and create .gitignore if missing =======
Public Sub InitializeProject()
    Dim fso As New Scripting.FileSystemObject
    Dim gitignorePath As String
    gitignorePath = joinPaths(Dirname, ".gitignore")

    ' Ensure repo root exists
    If Not fso.FolderExists(Dirname) Then
        fso.CreateFolder Dirname
    End If
    ' Ensure ./src exists
    If Not fso.FolderExists(SourceDirectory) Then
        fso.CreateFolder SourceDirectory
    End If

    ' Create a default .gitignore if missing
    If Not fso.FileExists(gitignorePath) Then
        With fso.OpenTextFile(gitignorePath, ForWriting, True)
            .WriteLine ("# Packages")
            .WriteLine ("node_modules")
            .WriteBlankLines 1
            .WriteLine ("# Word backup/lock files")
            .WriteLine ("~$*.do*")
            .WriteLine ("*.asd")
            .WriteLine ("*.wbk")
            .Close
        End With
    End If

    ' Initialize git in repo root (safe if already initialized)
    Bash script:="git init", keepCommandWindowOpen:=False
End Sub

' ======= Map a VBComponent to an export filename =======
Private Function getVBComponentFilename(ByRef component As VBComponent) As String
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_ClassModule
            getVBComponentFilename = component.name & ".cls"

        Case vbext_ComponentType.vbext_ct_StdModule
            getVBComponentFilename = component.name & ".bas"

        Case vbext_ComponentType.vbext_ct_MSForm
            getVBComponentFilename = component.name & ".frm"

        Case vbext_ComponentType.vbext_ct_Document
            ' Keep parity with your original: export as .cls
            getVBComponentFilename = component.name & ".cls"

        Case Else
            Debug.Print "Unknown component type: " & component.Type
            getVBComponentFilename = component.name & ".bas"
    End Select
End Function

' ======= Check if a component exists in current project by filename =======
Private Function componentExists(ByVal filename As String) As Boolean
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)

        If StrComp(getVBComponentFilename(component), filename, vbTextCompare) = 0 Then
            componentExists = True
            Exit Function
        End If
    Next index
End Function

' ======= Export all modules from Normal.dotm into ./src =======
Public Sub ExportComponentsToSourceFolder()
    Dim fso As New Scripting.FileSystemObject

    ' Warn if Normal.dotm not at specified path (still export from loaded Normal)
    If Len(Dir$(NORMAL_DOTM_PATH)) = 0 Then
        MsgBox "Warning: Normal.dotm not found at:" & vbCrLf & NORMAL_DOTM_PATH & vbCrLf & _
               "Proceeding to export from the Normal template currently loaded.", vbExclamation
    End If

    ' Ensure ./src exists; clear it
    If Not fso.FolderExists(SourceDirectory) Then
        fso.CreateFolder SourceDirectory
    Else
        Dim file As Scripting.file
        For Each file In fso.GetFolder(SourceDirectory).Files
            file.Delete True
        Next file
    End If

    ' Export each component
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)

        component.Export joinPaths(SourceDirectory, getVBComponentFilename(component))
        ' Forms will also emit a .frx alongside the .frm automatically
    Next index

    MsgBox "Export complete. Components saved to:" & vbCrLf & SourceDirectory, vbInformation
End Sub

' ======= Import from ./src into Normal.dotm (DESTRUCTIVE) =======
Public Sub DangerouslyImportComponentsFromSourceFolder()
    If MsgBox("Are you sure you want to import from the source folder (./src)? " & _
              "Existing modules may be removed. Continue?", vbYesNo + vbExclamation) = vbNo Then
        Exit Sub
    End If

    Dim fso As New Scripting.FileSystemObject
    If Not fso.FolderExists(SourceDirectory) Then
        MsgBox "Source folder not found: " & SourceDirectory, vbCritical
        Exit Sub
    End If

    ' Remove components that conflict by name (skip Project.bas and Document modules)
    Dim file As Scripting.file
    For Each file In fso.GetFolder(SourceDirectory).Files
        If LCase$(fso.GetExtensionName(file.name)) <> "frx" Then
            If componentExists(file.name) And LCase$(file.name) <> "project.bas" Then
                Dim component As VBComponent
                Set component = thisProjectsVBComponents.Item(fso.GetBaseName(file.name))

                If Not component Is Nothing Then
                    If component.Type <> vbext_ct_Document Then
                        thisProjectsVBComponents.Remove component
                    End If
                End If
            End If
        End If
    Next file

    ' Defer import to separate routine (original pattern preserved)
    Application.OnTime Now, "saftleyImportAfterCleanup"
End Sub

' ======= Perform the actual import after cleanup (original name preserved) =======
Private Sub saftleyImportAfterCleanup()
    Dim fso As New Scripting.FileSystemObject
    If Not fso.FolderExists(SourceDirectory) Then Exit Sub

    Dim file As Scripting.file
    For Each file In fso.GetFolder(SourceDirectory).Files
        If LCase$(fso.GetExtensionName(file.name)) <> "frx" Then
            If Not componentExists(file.name) Then
                thisProjectsVBComponents.Import joinPaths(SourceDirectory, file.name)
            End If
        End If
    Next file

    MsgBox "Import complete from: " & SourceDirectory, vbInformation
End Sub

' ======= Helpers for reporting =======
Private Function getVBComponentTypeName(ByRef component As VBComponent) As String
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_ClassModule: getVBComponentTypeName = "Class Module"
        Case vbext_ComponentType.vbext_ct_StdModule:   getVBComponentTypeName = "Module"
        Case vbext_ComponentType.vbext_ct_MSForm:      getVBComponentTypeName = "Form"
        Case vbext_ComponentType.vbext_ct_Document:    getVBComponentTypeName = "Document"
        Case Else:                                     getVBComponentTypeName = "Unknown"
    End Select
End Function

Private Function getComponentDetails(ByRef component As VBComponent) As String
    getComponentDetails = component.name & vbTab _
                          & getVBComponentTypeName(component) & vbTab _
                          & getVBComponentFilename(component)
End Function

Public Property Get ComponentsDetails() As String
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)

        ComponentsDetails = ComponentsDetails & getComponentDetails(component) & vbNewLine
    Next index
End Property

' ======= Dev helper: print filenames of current project components =======
Private Sub printDiffFromSourceFolder()
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)

        Debug.Print getVBComponentFilename(component)
    Next index
End Sub

' ======= Path join helper =======
Private Function joinPaths(ParamArray paths() As Variant) As String
    Dim fso As New Scripting.FileSystemObject
    Dim index As Long, acc As String, part As String
    For index = LBound(paths) To UBound(paths)
        part = Replace(CStr(paths(index)), "/", "\")
        If Len(acc) = 0 Then
            acc = part
        Else
            acc = fso.BuildPath(acc, part)
        End If
    Next
    joinPaths = acc
End Function


