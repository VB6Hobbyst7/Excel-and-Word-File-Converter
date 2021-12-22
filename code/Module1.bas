Attribute VB_Name = "Module1"

Function whichOption(frame As Variant, controlType As String) As Variant
Dim out As New Collection
    For Each Control In frame.Controls
        If UCase(TypeName(Control)) = UCase(controlType) Then
            If Control.Value = True Then
                out.Add Control
                If TypeName(frame) = "Frame" Then Exit For
            End If
        End If
    Next
    
    If out.Count = 1 Then
        Set whichOption = out(1)
    ElseIf out.Count > 1 Then
        Set whichOption = out
    End If
End Function


Public Function SelectFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Select a folder"
        .Show
        If .SelectedItems.Count > 0 Then
            SelectFolder = .SelectedItems.Item(1)
        Else
            'MsgBox "Folder is not selected."
        End If
    End With
End Function



Function IsFileOrFolder(Path) As String
    Dim retval
    retval = "Invalid"
    If (retval = "Invalid") And FileExists(Path) Then retval = "File"
    If (retval = "Invalid") And FolderExists(Path) Then retval = "Folder"
    IsFileOrFolder = retval
End Function

Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)
    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If
    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
    If Len(strFile) < 3 Then FileExists = False
End Function

Function FolderExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

